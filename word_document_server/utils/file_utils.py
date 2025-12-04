"""
File utility functions for Word Document Server.
"""
import os
import tempfile
from typing import Tuple, Optional, Callable, Any
from functools import wraps
import shutil

from word_document_server.utils.s3_utils import (
    is_s3_uri, download_from_s3, upload_to_s3, parse_s3_uri
)


def check_file_writeable(filepath: str) -> Tuple[bool, str]:
    """
    Check if a file can be written to.
    
    Args:
        filepath: Path to the file
        
    Returns:
        Tuple of (is_writeable, error_message)
    """
    # If file doesn't exist, check if directory is writeable
    if not os.path.exists(filepath):
        directory = os.path.dirname(filepath)
        # If no directory is specified (empty string), use current directory
        if directory == '':
            directory = '.'
        if not os.path.exists(directory):
            return False, f"Directory {directory} does not exist"
        if not os.access(directory, os.W_OK):
            return False, f"Directory {directory} is not writeable"
        return True, ""
    
    # If file exists, check if it's writeable
    if not os.access(filepath, os.W_OK):
        return False, f"File {filepath} is not writeable (permission denied)"
    
    # Try to open the file for writing to see if it's locked
    try:
        with open(filepath, 'a'):
            pass
        return True, ""
    except IOError as e:
        return False, f"File {filepath} is not writeable: {str(e)}"
    except Exception as e:
        return False, f"Unknown error checking file permissions: {str(e)}"


def create_document_copy(source_path: str, dest_path: Optional[str] = None) -> Tuple[bool, str, Optional[str]]:
    """
    Create a copy of a document.
    
    Args:
        source_path: Path to the source document
        dest_path: Optional path for the new document. If not provided, will use source_path + '_copy.docx'
        
    Returns:
        Tuple of (success, message, new_filepath)
    """
    if not os.path.exists(source_path):
        return False, f"Source document {source_path} does not exist", None
    
    if not dest_path:
        # Generate a new filename if not provided
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"
    
    try:
        # Simple file copy
        shutil.copy2(source_path, dest_path)
        return True, f"Document copied to {dest_path}", dest_path
    except Exception as e:
        return False, f"Failed to copy document: {str(e)}", None


def ensure_docx_extension(filename: str) -> str:
    """
    Ensure filename has .docx extension.

    Args:
        filename: The filename to check

    Returns:
        Filename with .docx extension
    """
    # For S3 URIs, handle the extension in the key part
    if is_s3_uri(filename):
        if not filename.endswith('.docx'):
            return filename + '.docx'
        return filename

    if not filename.endswith('.docx'):
        return filename + '.docx'
    return filename


class S3FileContext:
    """Context manager for transparent S3 file handling.

    Downloads S3 files to temp location, provides local paths for processing,
    and uploads results to a NEW S3 path (never overwrites original).

    IMPORTANT: For S3 files, edits are ALWAYS saved to a new file with timestamp.
    The original file is never modified. Use get_result_path() to get the new path.

    Example:
        with S3FileContext("s3://bucket/doc.docx") as ctx:
            doc = Document(ctx.local_path)
            # ... modify doc ...
            doc.save(ctx.local_path)
        # File is uploaded to NEW path: s3://bucket/doc_edited_1733318400.docx
        new_path = ctx.get_result_path()  # Returns the new S3 URI
    """

    def __init__(self, filename: str, read_only: bool = False, output_s3_uri: Optional[str] = None):
        """
        Args:
            filename: File path (local or S3 URI)
            read_only: If True, don't upload back to S3 on exit
            output_s3_uri: Optional explicit S3 URI for output (e.g., for conversions)
        """
        import time
        self.original_path = filename
        self.is_s3 = is_s3_uri(filename)
        self.read_only = read_only
        self.local_path = None
        self._temp_file = None
        self._output_temp_file = None
        self.output_local_path = None

        # For S3 write operations, generate a new output path (never overwrite)
        if output_s3_uri:
            self.output_s3_uri = output_s3_uri
        elif self.is_s3 and not read_only:
            # Generate new path with timestamp: doc.docx -> doc_edited_1733318400.docx
            self.output_s3_uri = self._generate_versioned_path(filename)
        else:
            self.output_s3_uri = None

    def _generate_versioned_path(self, s3_uri: str) -> str:
        """Generate a new S3 path with timestamp to avoid overwriting original."""
        import time
        base, ext = os.path.splitext(s3_uri)
        timestamp = int(time.time())
        return f"{base}_edited_{timestamp}{ext}"

    def __enter__(self):
        if self.is_s3:
            # Download from S3 to temp file
            success, message, local_path = download_from_s3(self.original_path)
            if not success:
                raise IOError(f"Failed to download from S3: {message}")
            self.local_path = local_path
            self._temp_file = local_path
        else:
            self.local_path = self.original_path

        # Handle output path for S3 writes or conversions
        if self.output_s3_uri:
            _, ext = os.path.splitext(self.output_s3_uri)
            fd, self.output_local_path = tempfile.mkstemp(suffix=ext if ext else '.docx')
            os.close(fd)
            self._output_temp_file = self.output_local_path
            # Copy input to output temp file so edits work on the copy
            if self.is_s3 and self.local_path:
                import shutil
                shutil.copy2(self.local_path, self.output_local_path)
                # Point local_path to output so saves go to the right place
                self.local_path = self.output_local_path
        else:
            self.output_local_path = self.local_path

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            # Upload to S3 if needed (and no exception occurred)
            if exc_type is None and not self.read_only:
                if self.output_s3_uri and os.path.exists(self.output_local_path):
                    # Upload to new S3 location (never overwrites original)
                    success, message = upload_to_s3(self.output_local_path, self.output_s3_uri)
                    if not success:
                        raise IOError(f"Failed to upload to S3: {message}")
        finally:
            # Clean up temp files
            if self._temp_file and os.path.exists(self._temp_file):
                try:
                    os.unlink(self._temp_file)
                except:
                    pass
            if self._output_temp_file and self._output_temp_file != self._temp_file and os.path.exists(self._output_temp_file):
                try:
                    os.unlink(self._output_temp_file)
                except:
                    pass

        return False  # Don't suppress exceptions

    def get_result_path(self) -> str:
        """Get the path to return to the user.

        For S3 files, this returns the NEW path where the edited file was saved.
        The original file remains untouched.
        """
        if self.output_s3_uri:
            return self.output_s3_uri
        return self.original_path


def resolve_file_for_read(filename: str) -> Tuple[str, bool, Optional[str]]:
    """Resolve a file path for reading, downloading from S3 if needed.

    Args:
        filename: File path (local or S3 URI)

    Returns:
        Tuple of (local_path, is_s3, original_s3_uri)

    Note: Caller is responsible for cleaning up temp files when is_s3=True
    """
    if is_s3_uri(filename):
        success, message, local_path = download_from_s3(filename)
        if not success:
            raise IOError(f"Failed to download from S3: {message}")
        return local_path, True, filename

    return filename, False, None


def check_file_exists(filename: str) -> Tuple[bool, str]:
    """Check if a file exists (supports both local and S3 paths).

    Args:
        filename: File path (local or S3 URI)

    Returns:
        Tuple of (exists, error_message)
    """
    if is_s3_uri(filename):
        # For S3, we can't easily check existence without downloading
        # Return True and let the download fail if it doesn't exist
        return True, ""

    if not os.path.exists(filename):
        return False, f"Document {filename} does not exist"

    return True, ""


def create_new_s3_document(s3_uri: str) -> Tuple[str, str]:
    """Create a temp file for a new S3 document.

    Args:
        s3_uri: S3 URI where the document will be uploaded

    Returns:
        Tuple of (local_temp_path, s3_uri)
    """
    _, ext = os.path.splitext(s3_uri)
    fd, local_path = tempfile.mkstemp(suffix=ext if ext else '.docx')
    os.close(fd)
    return local_path, s3_uri


def upload_if_s3(local_path: str, target_path: str) -> Tuple[bool, str]:
    """Upload a file to S3 if target is an S3 URI.

    Args:
        local_path: Path to local file
        target_path: Target path (local or S3 URI)

    Returns:
        Tuple of (success, message)
    """
    if is_s3_uri(target_path):
        return upload_to_s3(local_path, target_path)
    return True, ""


def cleanup_temp_file(filepath: str, is_temp: bool):
    """Clean up a temporary file if needed.

    Args:
        filepath: Path to the file
        is_temp: Whether the file is temporary and should be deleted
    """
    if is_temp and filepath and os.path.exists(filepath):
        try:
            os.unlink(filepath)
        except:
            pass
