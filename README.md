# Document to Quiz Excel Converter

A Node.js tool to convert quiz questions from Word documents into Excel format compatible with Quizziz.

## Installation

1. Clone or download this repository
2. Install dependencies:

```bash
npm install
```

The project uses the following dependencies:
- `mammoth` - For reading .docx files and preserving formatting
- `xlsx` - For creating Excel files

## Usage

1. Place your quiz document (`.docx` format) in the project directory
2. Update the file paths in `index.js` if needed:
   ```javascript
   const SOURCE_PATH = "quiz.docx"
   const OUTPUT_PATH = "quiz_export.xlsx"
   ```
3. Run the converter:
   ```bash
   npm start
   ```

The application will:
- Parse your quiz document
- Detect questions and options
- Identify correct answers based on bold formatting
- Generate an Excel file with the structured data

## Document Format Requirements

Your quiz document must follow this specific format:

### Structure
- **Every 5 lines** represent one complete question
- **Line 1**: Question text
- **Lines 2-5**: Four answer options

### Example:
```
What is the capital of France?
a) London
b) Paris
c) Berlin
d) Madrid
```

### Correct Answer Marking
- **Bold formatting** indicates the correct answer
- Only one option per question should be bold
- The application will detect `<strong>` or `<b>` tags in the document

## Output Format

The generated Excel file contains the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| Question Text | The question content | "What is the capital of France?" |
| Question Type | Type of question | "Multiple Choice" |
| Option 1 | First answer option | "London" |
| Option 2 | Second answer option | "Paris" |
| Option 3 | Third answer option | "Berlin" |
| Option 4 | Fourth answer option | "Madrid" |
| Option 5 | Reserved for future use | (empty) |
| Correct Answer | Index of correct option | "2" (for Option 2) |
| Time in seconds | Time limit for question | "30" |

### Notes:
- Correct answers are represented as **option numbers** (1, 2, 3, 4), not letters
- Option prefixes like "a)", "b)", etc. are automatically removed
- Default time limit is set to 30 seconds per question

## File Structure

```
├── index.js           # Main application file
├── package.json       # Node.js dependencies and scripts
├── quiz.docx         # Input quiz document (your file)
├── quiz_export.xlsx  # Output Excel file (generated)
└── README.md         # This documentation
```

## How It Works

1. **Document Reading**: Uses `mammoth` library to convert .docx to HTML while preserving formatting
2. **Content Parsing**: Splits content by paragraph tags and processes line by line
3. **Format Detection**: Groups every 5 lines as one question with 4 options
4. **Bold Detection**: Identifies bold formatting markers to find correct answers
5. **Data Structuring**: Organizes data into question objects with metadata
6. **Excel Generation**: Creates formatted Excel file using `xlsx` library

## Configuration

You can customize the following in `index.js`:

- **Source file path**: Change `SOURCE_PATH` variable
- **Output file path**: Change `OUTPUT_PATH` variable  
- **Question type**: Modify the default "Multiple Choice" value
- **Time limit**: Change the default "30" seconds value
- **Column formatting**: Adjust column widths in the `colWidths` array

## Troubleshooting

### Common Issues:

1. **"No correct answer detected"**
   - Ensure correct answers are formatted in bold in the Word document
   - Check that bold formatting is applied to the entire option text

2. **Questions not parsing correctly**
   - Verify your document follows the 5-line format (1 question + 4 options)
   - Remove extra blank lines between questions

3. **Option prefixes not removed**
   - The application automatically removes "a)", "b)", "c)", "d)" prefixes
   - Manual adjustment may be needed for other formats

## Requirements

- Node.js (version 14 or higher)
- npm (Node Package Manager)
- Input file must be in `.docx` format
- Document must follow the specified 5-line format

## License

This project is open source and available under the MIT License.

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

