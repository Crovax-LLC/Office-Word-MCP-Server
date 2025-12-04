"""
S3 utility functions for Word Document Server.

Provides transparent S3 file handling - download files from S3, process locally,
and upload results back to S3.
"""
import os
import tempfile
import logging
from typing import Tuple, Optional
from urllib.parse import urlparse
from contextlib import contextmanager

import boto3
from botocore.exceptions import ClientError, NoCredentialsError

logger = logging.getLogger(__name__)

# Initialize S3 client (uses IAM role credentials on EC2)
_s3_client = None


def get_s3_client():
    """Get or create S3 client."""
    global _s3_client
    if _s3_client is None:
        _s3_client = boto3.client('s3')
    return _s3_client


def is_s3_uri(path: str) -> bool:
    """Check if a path is an S3 URI.

    Args:
        path: File path or S3 URI

    Returns:
        True if path starts with s3://
    """
    return path.startswith('s3://')


def parse_s3_uri(s3_uri: str) -> Tuple[str, str]:
    """Parse an S3 URI into bucket and key.

    Args:
        s3_uri: S3 URI in format s3://bucket/key/path

    Returns:
        Tuple of (bucket_name, object_key)

    Raises:
        ValueError: If URI is not a valid S3 URI
    """
    if not is_s3_uri(s3_uri):
        raise ValueError(f"Not a valid S3 URI: {s3_uri}")

    parsed = urlparse(s3_uri)
    bucket = parsed.netloc
    key = parsed.path.lstrip('/')

    if not bucket:
        raise ValueError(f"No bucket specified in S3 URI: {s3_uri}")
    if not key:
        raise ValueError(f"No key specified in S3 URI: {s3_uri}")

    return bucket, key


def download_from_s3(s3_uri: str, local_path: Optional[str] = None) -> Tuple[bool, str, Optional[str]]:
    """Download a file from S3.

    Args:
        s3_uri: S3 URI of the file to download
        local_path: Optional local path to save the file. If not provided,
                   a temporary file will be created.

    Returns:
        Tuple of (success, message, local_file_path)
    """
    try:
        bucket, key = parse_s3_uri(s3_uri)
    except ValueError as e:
        return False, str(e), None

    try:
        s3 = get_s3_client()

        # Create local path if not provided
        if local_path is None:
            # Get the filename from the S3 key
            filename = os.path.basename(key)
            # Create temp file with same extension
            _, ext = os.path.splitext(filename)
            fd, local_path = tempfile.mkstemp(suffix=ext)
            os.close(fd)

        # Download the file
        s3.download_file(bucket, key, local_path)
        logger.info(f"Downloaded s3://{bucket}/{key} to {local_path}")

        return True, f"Downloaded from S3", local_path

    except NoCredentialsError:
        return False, "AWS credentials not configured", None
    except ClientError as e:
        error_code = e.response.get('Error', {}).get('Code', 'Unknown')
        if error_code == '404':
            return False, f"S3 object not found: {s3_uri}", None
        elif error_code == '403':
            return False, f"Access denied to S3 object: {s3_uri}", None
        else:
            return False, f"S3 error: {str(e)}", None
    except Exception as e:
        return False, f"Failed to download from S3: {str(e)}", None


def upload_to_s3(local_path: str, s3_uri: str) -> Tuple[bool, str]:
    """Upload a file to S3.

    Args:
        local_path: Path to the local file to upload
        s3_uri: S3 URI where the file should be uploaded

    Returns:
        Tuple of (success, message)
    """
    try:
        bucket, key = parse_s3_uri(s3_uri)
    except ValueError as e:
        return False, str(e)

    if not os.path.exists(local_path):
        return False, f"Local file does not exist: {local_path}"

    try:
        s3 = get_s3_client()

        # Determine content type based on extension
        content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        if local_path.endswith('.pdf'):
            content_type = 'application/pdf'

        # Upload the file
        s3.upload_file(
            local_path,
            bucket,
            key,
            ExtraArgs={'ContentType': content_type}
        )
        logger.info(f"Uploaded {local_path} to s3://{bucket}/{key}")

        return True, f"Uploaded to {s3_uri}"

    except NoCredentialsError:
        return False, "AWS credentials not configured"
    except ClientError as e:
        error_code = e.response.get('Error', {}).get('Code', 'Unknown')
        if error_code == '403':
            return False, f"Access denied uploading to S3: {s3_uri}"
        else:
            return False, f"S3 error: {str(e)}"
    except Exception as e:
        return False, f"Failed to upload to S3: {str(e)}"


def generate_output_s3_uri(input_s3_uri: str, suffix: str = "_output", new_extension: Optional[str] = None) -> str:
    """Generate an output S3 URI based on an input URI.

    Args:
        input_s3_uri: The input S3 URI
        suffix: Suffix to add before the extension (default: "_output")
        new_extension: New file extension (e.g., ".pdf"). If None, keeps original.

    Returns:
        New S3 URI for the output file
    """
    bucket, key = parse_s3_uri(input_s3_uri)

    base, ext = os.path.splitext(key)
    if new_extension:
        ext = new_extension

    new_key = f"{base}{suffix}{ext}"
    return f"s3://{bucket}/{new_key}"


@contextmanager
def s3_file_handler(input_path: str, output_path: Optional[str] = None, upload_output: bool = True):
    """Context manager for transparent S3 file handling.

    Downloads S3 files to temp location, yields local paths for processing,
    and uploads results back to S3 on exit.

    Args:
        input_path: Input file path (local or S3 URI)
        output_path: Output file path (local or S3 URI). If None and input is S3,
                    generates an output path in the same S3 location.
        upload_output: Whether to upload the output file to S3 (default: True)

    Yields:
        Tuple of (local_input_path, local_output_path, is_s3_input, output_s3_uri)

    Example:
        with s3_file_handler("s3://bucket/input.docx") as (local_in, local_out, is_s3, s3_out):
            # Process local_in, write to local_out
            doc = Document(local_in)
            doc.save(local_out)
        # File is automatically uploaded to S3 on exit
    """
    is_s3_input = is_s3_uri(input_path)
    is_s3_output = output_path and is_s3_uri(output_path)

    local_input = None
    local_output = None
    output_s3_uri = None
    temp_files = []

    try:
        # Handle input
        if is_s3_input:
            success, message, local_input = download_from_s3(input_path)
            if not success:
                raise IOError(f"Failed to download input from S3: {message}")
            temp_files.append(local_input)
        else:
            local_input = input_path

        # Handle output path
        if output_path:
            if is_s3_output:
                # Create temp file for output
                _, ext = os.path.splitext(output_path)
                fd, local_output = tempfile.mkstemp(suffix=ext if ext else '.docx')
                os.close(fd)
                temp_files.append(local_output)
                output_s3_uri = output_path
            else:
                local_output = output_path
        elif is_s3_input:
            # Generate output path in same S3 location
            output_s3_uri = generate_output_s3_uri(input_path, suffix="")
            _, ext = os.path.splitext(output_s3_uri)
            fd, local_output = tempfile.mkstemp(suffix=ext if ext else '.docx')
            os.close(fd)
            temp_files.append(local_output)
        else:
            # Local input, no output specified - use input path as output
            local_output = local_input

        yield local_input, local_output, is_s3_input, output_s3_uri

        # Upload output to S3 if needed
        if upload_output and output_s3_uri and os.path.exists(local_output):
            success, message = upload_to_s3(local_output, output_s3_uri)
            if not success:
                raise IOError(f"Failed to upload output to S3: {message}")

    finally:
        # Clean up temp files
        for temp_file in temp_files:
            try:
                if temp_file and os.path.exists(temp_file):
                    os.unlink(temp_file)
            except Exception as e:
                logger.warning(f"Failed to clean up temp file {temp_file}: {e}")


def resolve_s3_path(path: str) -> Tuple[str, bool, Optional[str]]:
    """Resolve a path that might be an S3 URI to a local path.

    Downloads from S3 if needed, returns local path for processing.

    Args:
        path: File path (local or S3 URI)

    Returns:
        Tuple of (local_path, is_s3, original_s3_uri)
        - local_path: Path to use for local processing
        - is_s3: Whether the original path was an S3 URI
        - original_s3_uri: The original S3 URI if is_s3 is True, else None
    """
    if is_s3_uri(path):
        success, message, local_path = download_from_s3(path)
        if not success:
            raise IOError(f"Failed to download from S3: {message}")
        return local_path, True, path
    else:
        return path, False, None
