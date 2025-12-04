# Word MCP Server - Laravel Backend Integration Guide

## Overview

The Word MCP Server is an internal service that provides Microsoft Word document manipulation capabilities via the Model Context Protocol (MCP). It enables AI agents and backend services to create, edit, and format Word documents programmatically.

## Connection Details

| Property | Value |
|----------|-------|
| **Endpoint** | `http://172.31.29.248:8000/mcp` |
| **Transport** | `streamable-http` |
| **Protocol** | MCP (Model Context Protocol) |
| **Region** | eu-central-1 |
| **Access** | Internal VPC only (BauGpt services) |

## Installation

Install the PHP MCP client via Composer:

```bash
composer require php-mcp/client
```

ReactPHP dependencies will be installed automatically.

## Configuration

### Create a Service Class

Create a service class to manage the Word MCP connection:

```php
<?php
// app/Services/WordMcpService.php

namespace App\Services;

use PhpMcp\Client\Client;
use PhpMcp\Client\Enum\TransportType;
use PhpMcp\Client\Model\Capabilities as ClientCapabilities;
use PhpMcp\Client\ServerConfig;
use Illuminate\Support\Facades\Log;

class WordMcpService
{
    private ?Client $client = null;
    private string $serverUrl;

    public function __construct()
    {
        $this->serverUrl = config('services.word_mcp.url', 'http://172.31.29.248:8000/mcp');
    }

    /**
     * Get or create the MCP client connection
     */
    public function getClient(): Client
    {
        if ($this->client === null || !$this->client->isReady()) {
            $this->connect();
        }

        return $this->client;
    }

    /**
     * Connect to the Word MCP Server
     */
    public function connect(): void
    {
        $config = new ServerConfig(
            name: 'word_mcp_server',
            transport: TransportType::Http,
            timeout: 60.0,
            url: $this->serverUrl,
        );

        $this->client = Client::make()
            ->withClientInfo('BauGpt-Laravel', '1.0')
            ->withCapabilities(ClientCapabilities::forClient())
            ->withServerConfig($config)
            ->build();

        $this->client->initialize();

        Log::info('Connected to Word MCP Server', ['url' => $this->serverUrl]);
    }

    /**
     * Disconnect from the server
     */
    public function disconnect(): void
    {
        if ($this->client !== null) {
            $this->client->disconnect();
            $this->client = null;
        }
    }

    /**
     * Call a tool on the MCP server
     */
    public function callTool(string $toolName, array $arguments = []): mixed
    {
        try {
            $client = $this->getClient();
            return $client->callTool($toolName, $arguments);
        } catch (\Throwable $e) {
            Log::error('Word MCP tool call failed', [
                'tool' => $toolName,
                'arguments' => $arguments,
                'error' => $e->getMessage(),
            ]);
            throw $e;
        }
    }

    /**
     * List all available tools
     */
    public function listTools(): array
    {
        $client = $this->getClient();
        return $client->listTools();
    }
}
```

### Register the Service Provider

Add to `config/services.php`:

```php
'word_mcp' => [
    'url' => env('WORD_MCP_URL', 'http://172.31.29.248:8000/mcp'),
],
```

Add to your `.env`:

```env
WORD_MCP_URL=http://172.31.29.248:8000/mcp
```

Register as a singleton in `app/Providers/AppServiceProvider.php`:

```php
use App\Services\WordMcpService;

public function register(): void
{
    $this->app->singleton(WordMcpService::class, function ($app) {
        return new WordMcpService();
    });
}
```

## S3 Integration

The Word MCP Server supports direct S3 file paths. You can pass S3 URIs anywhere a filename is expected, and the server will automatically:
1. Download the file from S3
2. Process it locally
3. Upload the result back to S3

### S3 URI Format

```
s3://bucket-name/path/to/file.docx
```

### Examples

```php
// Create a document directly in S3
$this->wordService->callTool('create_document', [
    'filename' => 's3://baugpt-documents/reports/monthly_report.docx',
    'title' => 'Monthly Report',
    'author' => 'BauGpt System',
]);

// Read an existing document from S3
$text = $this->wordService->callTool('get_document_text', [
    'filename' => 's3://baugpt-documents/templates/contract.docx',
]);

// Modify a document in S3 (downloads, modifies, uploads back)
$this->wordService->callTool('search_and_replace', [
    'filename' => 's3://baugpt-documents/contracts/contract_001.docx',
    'find_text' => '{{CLIENT_NAME}}',
    'replace_text' => 'Acme Corp',
]);

// Copy from S3 to S3
$this->wordService->callTool('copy_document', [
    'source_filename' => 's3://baugpt-documents/templates/invoice.docx',
    'destination_filename' => 's3://baugpt-documents/invoices/invoice_2024_001.docx',
]);

// Convert S3 document to PDF (result uploaded to S3)
$this->wordService->callTool('convert_to_pdf', [
    'filename' => 's3://baugpt-documents/reports/report.docx',
    'output_filename' => 's3://baugpt-documents/reports/report.pdf',
]);

// Add image from S3 to document in S3
$this->wordService->callTool('add_picture', [
    'filename' => 's3://baugpt-documents/reports/report.docx',
    'image_path' => 's3://baugpt-assets/logos/company_logo.png',
    'width' => 2.0,
]);
```

### Mixed Local and S3 Paths

You can mix local and S3 paths:

```php
// Copy local template to S3
$this->wordService->callTool('copy_document', [
    'source_filename' => '/tmp/template.docx',
    'destination_filename' => 's3://baugpt-documents/output/report.docx',
]);

// Download from S3 to local for processing
$this->wordService->callTool('copy_document', [
    'source_filename' => 's3://baugpt-documents/templates/base.docx',
    'destination_filename' => '/tmp/working_copy.docx',
]);
```

### Supported Tools for S3

All document tools support S3 URIs, including:
- `create_document`
- `get_document_info`, `get_document_text`, `get_document_outline`
- `copy_document`, `merge_documents`
- `add_heading`, `add_paragraph`, `add_table`, `add_picture`
- `search_and_replace`
- `convert_to_pdf`
- `format_text`, `format_table`, table formatting tools
- And all other document manipulation tools

### S3 Permissions

The EC2 instance has S3 full access via its IAM role (`WordMcpServerEC2Role`). Ensure your S3 bucket allows access from the EC2 instance.

---

## Usage Examples

### Basic Document Creation

```php
<?php

use App\Services\WordMcpService;

class ReportController extends Controller
{
    public function __construct(
        private WordMcpService $wordService
    ) {}

    public function generateReport()
    {
        $filename = '/tmp/report_' . time() . '.docx';

        // Create document
        $this->wordService->callTool('create_document', [
            'filename' => $filename,
            'title' => 'Monthly Report',
            'author' => 'BauGpt System',
        ]);

        // Add heading
        $this->wordService->callTool('add_heading', [
            'filename' => $filename,
            'text' => 'Monthly Performance Report',
            'level' => 1,
            'font_name' => 'Arial',
            'font_size' => 24,
            'bold' => true,
        ]);

        // Add paragraph
        $this->wordService->callTool('add_paragraph', [
            'filename' => $filename,
            'text' => 'This report summarizes the key metrics for the month.',
            'font_name' => 'Arial',
            'font_size' => 12,
        ]);

        return response()->json([
            'success' => true,
            'filename' => $filename,
        ]);
    }
}
```

### Creating Tables with Formatting

```php
public function generateTableReport()
{
    $filename = '/tmp/table_report_' . time() . '.docx';

    // Create document
    $this->wordService->callTool('create_document', [
        'filename' => $filename,
        'title' => 'Sales Report',
    ]);

    // Add heading
    $this->wordService->callTool('add_heading', [
        'filename' => $filename,
        'text' => 'Q4 Sales Summary',
        'level' => 1,
    ]);

    // Add table
    $this->wordService->callTool('add_table', [
        'filename' => $filename,
        'rows' => 5,
        'cols' => 4,
        'data' => [
            ['Product', 'Q1', 'Q2', 'Total'],
            ['Widget A', '€10,000', '€12,000', '€22,000'],
            ['Widget B', '€8,000', '€9,500', '€17,500'],
            ['Widget C', '€15,000', '€18,000', '€33,000'],
            ['Total', '€33,000', '€39,500', '€72,500'],
        ],
    ]);

    // Format header row (blue background, white text)
    $this->wordService->callTool('highlight_table_header', [
        'filename' => $filename,
        'table_index' => 0,
        'header_color' => '2E86AB',
        'text_color' => 'FFFFFF',
    ]);

    // Add alternating row colors
    $this->wordService->callTool('apply_table_alternating_rows', [
        'filename' => $filename,
        'table_index' => 0,
        'color1' => 'FFFFFF',
        'color2' => 'F5F5F5',
    ]);

    // Auto-fit columns
    $this->wordService->callTool('auto_fit_table_columns', [
        'filename' => $filename,
        'table_index' => 0,
    ]);

    return $filename;
}
```

### Search and Replace

```php
public function updateDocument(string $filename, array $replacements)
{
    foreach ($replacements as $find => $replace) {
        $this->wordService->callTool('search_and_replace', [
            'filename' => $filename,
            'find_text' => $find,
            'replace_text' => $replace,
        ]);
    }

    return true;
}

// Usage:
$this->updateDocument('/tmp/template.docx', [
    '{{CLIENT_NAME}}' => 'Acme Corp',
    '{{DATE}}' => now()->format('d.m.Y'),
    '{{AMOUNT}}' => '€25,000',
]);
```

### Converting to PDF

```php
public function convertToPdf(string $docxPath): string
{
    $pdfPath = str_replace('.docx', '.pdf', $docxPath);

    $result = $this->wordService->callTool('convert_to_pdf', [
        'filename' => $docxPath,
        'output_filename' => $pdfPath,
    ]);

    return $pdfPath;
}
```

### Adding Lists

```php
public function addBulletList(string $filename, array $items, string $afterText)
{
    $this->wordService->callTool('insert_numbered_list_near_text', [
        'filename' => $filename,
        'target_text' => $afterText,
        'list_items' => $items,
        'position' => 'after',
        'bullet_type' => 'bullet',  // or 'number' for numbered list
    ]);
}

// Usage:
$this->addBulletList($filename, [
    'First action item',
    'Second action item',
    'Third action item',
], 'Action Items:');
```

### Working with Footnotes

```php
public function addFootnote(string $filename, string $searchText, string $footnoteText)
{
    $this->wordService->callTool('add_footnote_robust', [
        'filename' => $filename,
        'search_text' => $searchText,
        'footnote_text' => $footnoteText,
        'validate_location' => true,
    ]);
}
```

### Extracting Document Content

```php
public function getDocumentText(string $filename): string
{
    $result = $this->wordService->callTool('get_document_text', [
        'filename' => $filename,
    ]);

    return $result;
}

public function getDocumentInfo(string $filename): array
{
    $result = $this->wordService->callTool('get_document_info', [
        'filename' => $filename,
    ]);

    return $result;
}
```

## Complete Laravel Job Example

```php
<?php
// app/Jobs/GenerateMonthlyReport.php

namespace App\Jobs;

use App\Services\WordMcpService;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;
use Illuminate\Support\Facades\Storage;

class GenerateMonthlyReport implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    public function __construct(
        private array $reportData,
        private string $userId,
    ) {}

    public function handle(WordMcpService $wordService): void
    {
        $filename = '/tmp/monthly_report_' . $this->userId . '_' . time() . '.docx';

        try {
            // Create document
            $wordService->callTool('create_document', [
                'filename' => $filename,
                'title' => 'Monthly Report - ' . now()->format('F Y'),
                'author' => 'BauGpt Reporting System',
            ]);

            // Title
            $wordService->callTool('add_heading', [
                'filename' => $filename,
                'text' => $this->reportData['title'],
                'level' => 1,
                'font_name' => 'Arial',
                'bold' => true,
            ]);

            // Summary section
            $wordService->callTool('add_heading', [
                'filename' => $filename,
                'text' => 'Executive Summary',
                'level' => 2,
                'border_bottom' => true,
            ]);

            $wordService->callTool('add_paragraph', [
                'filename' => $filename,
                'text' => $this->reportData['summary'],
            ]);

            // Metrics table
            if (!empty($this->reportData['metrics'])) {
                $wordService->callTool('add_heading', [
                    'filename' => $filename,
                    'text' => 'Key Metrics',
                    'level' => 2,
                ]);

                $tableData = array_merge(
                    [['Metric', 'Value', 'Change']],
                    $this->reportData['metrics']
                );

                $wordService->callTool('add_table', [
                    'filename' => $filename,
                    'rows' => count($tableData),
                    'cols' => 3,
                    'data' => $tableData,
                ]);

                $wordService->callTool('highlight_table_header', [
                    'filename' => $filename,
                    'table_index' => 0,
                    'header_color' => '1a5f7a',
                    'text_color' => 'FFFFFF',
                ]);
            }

            // Convert to PDF
            $pdfFilename = str_replace('.docx', '.pdf', $filename);
            $wordService->callTool('convert_to_pdf', [
                'filename' => $filename,
                'output_filename' => $pdfFilename,
            ]);

            // Upload to S3
            Storage::disk('s3')->put(
                "reports/{$this->userId}/" . basename($pdfFilename),
                file_get_contents($pdfFilename)
            );

            // Cleanup temp files
            @unlink($filename);
            @unlink($pdfFilename);

        } catch (\Throwable $e) {
            \Log::error('Report generation failed', [
                'user_id' => $this->userId,
                'error' => $e->getMessage(),
            ]);
            throw $e;
        } finally {
            $wordService->disconnect();
        }
    }
}
```

## Available Tools Reference

### Document Management

| Tool | Description | Parameters |
|------|-------------|------------|
| `create_document` | Create new Word document | `filename`, `title?`, `author?` |
| `copy_document` | Copy a document | `source_filename`, `destination_filename?` |
| `get_document_info` | Get document metadata | `filename` |
| `get_document_text` | Extract all text | `filename` |
| `get_document_outline` | Get structure/headings | `filename` |
| `list_available_documents` | List .docx files | `directory` (default: ".") |
| `convert_to_pdf` | Convert to PDF | `filename`, `output_filename?` |

### Content Creation

| Tool | Parameters |
|------|------------|
| `add_paragraph` | `filename`, `text`, `style?`, `font_name?`, `font_size?`, `bold?`, `italic?`, `color?` |
| `add_heading` | `filename`, `text`, `level` (1-9), `font_name?`, `font_size?`, `bold?`, `italic?`, `border_bottom?` |
| `add_table` | `filename`, `rows`, `cols`, `data?` (array of arrays) |
| `add_picture` | `filename`, `image_path`, `width?` |
| `add_page_break` | `filename` |
| `delete_paragraph` | `filename`, `paragraph_index` (0-based) |
| `search_and_replace` | `filename`, `find_text`, `replace_text` |

### Content Insertion (Relative Positioning)

| Tool | Parameters |
|------|------------|
| `insert_header_near_text` | `filename`, `target_text` OR `target_paragraph_index`, `header_title`, `position` ("before"/"after"), `header_style?` |
| `insert_line_or_paragraph_near_text` | `filename`, `target_text` OR `target_paragraph_index`, `line_text`, `position`, `line_style?` |
| `insert_numbered_list_near_text` | `filename`, `target_text` OR `target_paragraph_index`, `list_items` (array), `position`, `bullet_type` ("bullet"/"number") |

### Text Formatting

| Tool | Parameters |
|------|------------|
| `format_text` | `filename`, `paragraph_index`, `start_pos`, `end_pos`, `bold?`, `italic?`, `underline?`, `color?`, `font_size?`, `font_name?` |
| `create_custom_style` | `filename`, `style_name`, `bold?`, `italic?`, `font_size?`, `font_name?`, `color?`, `base_style?` |

### Table Formatting

| Tool | Parameters |
|------|------------|
| `format_table` | `filename`, `table_index`, `has_header_row?`, `border_style?`, `shading?` |
| `set_table_cell_shading` | `filename`, `table_index`, `row_index`, `col_index`, `fill_color`, `pattern?` |
| `apply_table_alternating_rows` | `filename`, `table_index`, `color1?`, `color2?` |
| `highlight_table_header` | `filename`, `table_index`, `header_color?`, `text_color?` |
| `merge_table_cells` | `filename`, `table_index`, `start_row`, `start_col`, `end_row`, `end_col` |
| `merge_table_cells_horizontal` | `filename`, `table_index`, `row_index`, `start_col`, `end_col` |
| `merge_table_cells_vertical` | `filename`, `table_index`, `col_index`, `start_row`, `end_row` |
| `set_table_cell_alignment` | `filename`, `table_index`, `row_index`, `col_index`, `horizontal?`, `vertical?` |
| `format_table_cell_text` | `filename`, `table_index`, `row_index`, `col_index`, `text_content?`, `bold?`, `italic?`, `underline?`, `color?`, `font_size?`, `font_name?` |
| `set_table_cell_padding` | `filename`, `table_index`, `row_index`, `col_index`, `top?`, `bottom?`, `left?`, `right?`, `unit?` |
| `set_table_column_width` | `filename`, `table_index`, `col_index`, `width`, `width_type?` |
| `auto_fit_table_columns` | `filename`, `table_index` |

### Footnotes & Endnotes

| Tool | Parameters |
|------|------------|
| `add_footnote_robust` | `filename`, `search_text?`, `paragraph_index?`, `footnote_text`, `validate_location?`, `auto_repair?` |
| `add_endnote_to_document` | `filename`, `paragraph_index`, `endnote_text` |
| `delete_footnote_robust` | `filename`, `footnote_id?`, `search_text?`, `clean_orphans?` |

### Comments

| Tool | Parameters |
|------|------------|
| `get_all_comments` | `filename` |
| `get_comments_by_author` | `filename`, `author` |
| `get_comments_for_paragraph` | `filename`, `paragraph_index` |

### Document Protection

| Tool | Parameters |
|------|------------|
| `protect_document` | `filename`, `password` |
| `unprotect_document` | `filename`, `password` |

### Search

| Tool | Parameters |
|------|------------|
| `get_paragraph_text_from_document` | `filename`, `paragraph_index` |
| `find_text_in_document` | `filename`, `text_to_find`, `match_case?`, `whole_word?` |

## Important Conventions

### Indexing
All indices are **0-based**:
- First paragraph = index `0`
- First table = index `0`
- First row/column = index `0`

### Colors
Use **hex RGB without `#` prefix**:
```php
'color' => 'FF0000'   // Red
'color' => '00FF00'   // Green
'color' => '2E86AB'   // Blue
```

### File Paths
Two formats are supported:
- **Local paths**: Absolute paths on EC2 filesystem (e.g., `/tmp/document.docx`)
- **S3 URIs**: Direct S3 paths (e.g., `s3://bucket/path/document.docx`)

**Recommended workflow**:
- Use S3 URIs directly - the server handles download/upload automatically
- For temporary work: Use `/tmp/` and upload to S3 when done

### Units
- Font sizes: **points** (e.g., 12, 14, 24)
- Width types: `"points"` or `"percentage"`
- Padding: `"points"` or `"percent"`

## Error Handling

```php
use App\Services\WordMcpService;

class DocumentService
{
    public function __construct(private WordMcpService $wordService) {}

    public function safeCallTool(string $tool, array $args): array
    {
        try {
            $result = $this->wordService->callTool($tool, $args);
            return [
                'success' => true,
                'result' => $result,
            ];
        } catch (\Throwable $e) {
            \Log::error("Word MCP Error: {$tool}", [
                'arguments' => $args,
                'error' => $e->getMessage(),
                'trace' => $e->getTraceAsString(),
            ]);

            return [
                'success' => false,
                'error' => $e->getMessage(),
            ];
        }
    }
}
```

## Debugging

### List Available Tools

```php
$tools = $this->wordService->listTools();
foreach ($tools as $tool) {
    dump($tool->name, $tool->description);
}
```

### Check Server Logs

```bash
# Connect to EC2 via SSM
aws --profile baugpt ssm start-session --target i-0ef6d1d79126734d0

# View container logs
docker logs word-mcp-server -f --tail 100
```

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| Connection timeout | Server unreachable | Check security groups, verify VPC connectivity |
| Tool not found | Typo in tool name | Use `listTools()` to verify exact names |
| File not found | Invalid path | Use absolute paths starting with `/` |
| Index out of range | Invalid paragraph/table index | Use `get_document_info` first |

## Support

- **Server Instance**: `i-0ef6d1d79126734d0`
- **Private IP**: `172.31.29.248`
- **AWS Region**: eu-central-1

## References

- [php-mcp/client GitHub](https://github.com/php-mcp/client) - PHP MCP Client Library
- [Model Context Protocol Specification](https://modelcontextprotocol.io/)
