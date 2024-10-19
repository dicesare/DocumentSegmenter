# Document Segmenter

This project provides a Python script that segments documents (PDF, DOCX, TXT) into different components such as titles, subtitles, and paragraphs. The segments are then saved in a CSV file for further analysis or processing.

## Features

- Supports segmentation of **DOCX**, **PDF**, and **TXT** files.
- Uses `pymupdf4llm` to convert PDF files into Markdown before segmentation.
- Identifies **titles**, **subtitles**, and **paragraphs** based on document styles (e.g., heading levels in DOCX).
- Saves segmented content into a **CSV file** with the segment type, importance, and text.

## Installation

### Requirements
- Python 3.7+
- Install the necessary libraries:
    ```bash
    pip install docx pymupdf4llm
    ```

## Usage

1. Clone the repository or download the Python script.
2. Prepare the document (PDF, DOCX, or TXT) you want to segment.
3. Replace the file path in the `file_path` variable with the path to your document.
4. Run the script to generate the segmented CSV output.

```bash
python script_name.py
```
## Example Usage

```python
# Example of using the DocumentSegmenter to process a file
file_path = r'data/original/your_document_to_chunk.pdf'  # Replace with your document path
output_dir = 'data/chunk'  # Folder to save the CSV file
segmenter = DocumentSegmenter(file_path, output_dir)
segmenter.process()
```

2. **Segmentation**: The script identifies titles, subtitles, and paragraphs, creating segments with associated importance.

3. **Saving**: The script saves the segmented text into a CSV file, where each row contains:
    - **Segment Type**: Whether it is a title, subtitle, or paragraph.
    - **Importance**: The importance level (e.g., 2.0 for titles, 1.8 for subtitles).
    - **Text**: The actual text of the segment.

## Example Output

The output CSV file will contain the following columns:

| Segment Type | Importance | Text                             |
|--------------|------------|----------------------------------|
| title        | 2.0        | Main Title of the Document       |
| paragraph    | 1.0        | Content of the first paragraph   |
| subtitle     | 1.8        | Subtitle for Section 1           |
| paragraph    | 1.0        | Content of the paragraph under Section 1 |

## Contribution

Feel free to submit issues or pull requests for any improvements or additional features.

## License

This project is open source and available under the MIT License.