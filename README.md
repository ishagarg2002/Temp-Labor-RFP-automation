# Temp Labor RFP Generator

AI-powered tool that reads any client document (RFI, meeting notes, interview transcripts, prior RFPs) and generates a complete Temporary Labor RFP as an Excel workbook.

## How It Works

1. Upload/provide any client input file (PDF, Word, Excel, Text, CSV)
2. AI reads the file, extracts client info, and generates a full RFP
3. Output is an Excel file with the same tab structure and formatting as the reference template

## Output Structure (9 tabs)

| Tab | Content |
|-----|---------|
| 1. Cover Letter | Client info, objectives, scope, timeline, contact |
| 2. Contents & Response Summary | Workbook index |
| 3. Requirements | Business requirements for staffing agencies |
| 4. Questionnaire | Capability questions for suppliers |
| 5. Service Level Agreements | KPIs with targets and measurement criteria |
| 6a-d. Bid Sheets | Roles by region (APAC, EMEA, LATAM, NA) with pay rates |
| 7. Bid Sheet Mark-Up | Mark-up structure by location and job family |
| 8. Discounts | Conversion fees, volume rebates, tenure discounts |
| 9. Conversion Rates | Currency conversion reference table |

## Setup

```bash
# Install dependencies
pip install -r requirements.txt

# Create .env file with your API key
echo ANTHROPIC_API_KEY=your-key-here > .env

# Place your reference template Excel in the project folder
# Update TEMPLATE path in run_rfp.py if needed
```

## Usage

```bash
python run_rfp.py "path/to/client_file.xlsx"
```

Output Excel appears in `./output/` and opens automatically.

## Dashboard (optional)

Open `rfp_builder.html` in a browser for a visual interface to:
- Upload input files
- Get the run command
- Preview generated output

## Files

| File | Purpose |
|------|---------|
| `run_rfp.py` | Main script — reads input, calls AI, generates Excel |
| `rfp_builder.html` | Browser dashboard UI |
| `rfp_dashboard.html` | Alternative form-based UI |
| `.env` | API key (not committed to git) |
| `requirements.txt` | Python dependencies |

## Notes

- The reference template is used for formatting/structure only — no sample data carries over
- All logos and branding are removed from output
- Works with any client — just provide their info as input
- API calls go through the configured proxy endpoint
