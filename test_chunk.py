import os
import csv
from docx import Document
from datetime import datetime
import pymupdf4llm as pymu  # Library to handle PDF extraction


class Segment:
    """
    Represents a segment of extracted text (title, subtitle, paragraph).

    :param text: The segment's text content
    :param segment_type: Type of the segment (title, subtitle, paragraph)
    :param importance: Relative importance of the segment (default is 1.0)
    :return: A Segment object
    """

    def __init__(self, text, segment_type, importance=1.0):
        self.text = text
        self.segment_type = segment_type
        self.importance = importance

    def __repr__(self):
        return f"Segment(type={self.segment_type}, importance={self.importance}, text={self.text[:30]}...)"


class DocumentSegmenter:
    """
    A class that segments and extracts titles, subtitles, and paragraphs from DOCX, PDF, and TXT documents.

    :param file_path: The path to the file to be processed
    :param output_dir: Directory where the CSV files will be saved
    :return: A DocumentSegmenter object

    EXAMPLE USAGE:
    file_path = r'data/original/your_document_to_chunk.pdf' or document.docx  # Replace with your file path
    output_dir = 'data/chunk'  # Folder to save the CSV file; the folder will be created if it doesn't exist
    """

    def __init__(self, file_path, output_dir):
        self.file_path = file_path
        self.segments = []
        self.file_extension = os.path.splitext(file_path)[1].lower()
        self.output_dir = output_dir

        # Ensure the output directory exists, otherwise create it
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            print(f"Output directory created: {self.output_dir}")
        else:
            print(f"Output directory already exists: {self.output_dir}")

    def load_document(self):
        """
        Loads the document based on its file type (DOCX, PDF, TXT) and returns its content as plain text or Markdown.

        :return: Plain text or Markdown representation of the document
        """
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"The file {self.file_path} does not exist.")

        if self.file_extension == '.docx':
            return self._load_docx()
        elif self.file_extension == '.pdf':
            return self._load_pdf()
        elif self.file_extension == '.txt':
            return self._load_txt()
        else:
            raise ValueError(f"Unsupported file format: {self.file_extension}")

    def _load_docx(self):
        """
        Loads a DOCX file and returns its content as plain text.

        :return: The plain text content of the document
        """
        doc = Document(self.file_path)
        # Convert all paragraphs into a single string joined by newlines
        return "\n".join([para.text for para in doc.paragraphs])

    def _load_pdf(self):
        """
        Uses pymupdf4llm to extract text from a PDF file and convert it to Markdown.

        :return: The converted Markdown text from the PDF
        """
        try:
            return pymu.to_markdown(self.file_path)  # Converts the PDF content to Markdown
        except Exception as e:
            print(f"Error reading the PDF file: {e}")
            return None

    def _load_txt(self):
        """
        Reads a TXT file and returns its content as plain text.

        :return: The plain text content of the TXT file
        """
        with open(self.file_path, 'r', encoding='utf-8') as file:
            return file.read()

    def segment_document(self):
        """
        Segments the loaded document based on text styles and content type (title, subtitle, paragraph).
        The segmentation is determined differently for DOCX, PDF, and TXT.
        """
        text = self.load_document()
        if text is None:
            raise ValueError("Unable to load the document content.")

        # Call appropriate segmentation method based on file extension
        if self.file_extension == '.docx':
            self._segment_docx(text)
        else:
            self._segment_plain_text(text)  # Default segmentation for PDF and TXT

    def _segment_docx(self, text):
        """
        Segments a DOCX document by detecting titles, subtitles, and paragraphs based on text styles (e.g., Heading 1, Heading 2).

        This method ensures that titles and subtitles are properly identified based on their style level in DOCX.

        :return: A list of detected segments (title, subtitle, paragraph)
        """
        doc = Document(self.file_path)
        current_text = ""
        current_segment_type = "paragraph"
        current_importance = 1.0

        # Iterate through each paragraph in the DOCX file
        for para in doc.paragraphs:
            # Replace non-breaking spaces and strip excess whitespace
            line = para.text.replace('\u00A0', ' ').strip()

            if not line:
                continue  # Skip empty paragraphs

            # Detect titles and subtitles based on heading levels
            if para.style.name.startswith('Heading'):  # Title or subtitle
                if current_text and current_segment_type == 'paragraph':
                    # Save the current paragraph before switching segment type
                    self.segments.append(Segment(current_text, current_segment_type, current_importance))

                # Determine if it's a title or subtitle based on heading level
                level = int(para.style.name.split()[1])
                current_segment_type = 'title' if level == 1 else 'subtitle'
                current_importance = 2.0 if level == 1 else 1.8
                self.segments.append(Segment(line, current_segment_type, current_importance))

                # Reset for the next paragraph
                current_text = ""
                current_segment_type = "paragraph"
                current_importance = 1.0
            else:
                # Concatenate the line to the current paragraph
                current_text += line + " "

        # Add any remaining paragraph text to the segments
        if current_text:
            self.segments.append(Segment(current_text.strip(), current_segment_type, current_importance))

    def _segment_plain_text(self, text):
        """
        Segments plain text by splitting it into paragraphs.
        This method is also used for PDFs that are converted to Markdown.

        :param text: The plain text content to be segmented
        """
        paragraphs = text.split("\n")
        for paragraph in paragraphs:
            if paragraph.strip():  # Skip empty lines
                self.segments.append(Segment(paragraph.strip(), "paragraph", 1.0))

    def save_segments_to_csv(self, output_csv):
        """
        Saves the detected segments into a CSV file.

        :param output_csv: The name of the output CSV file
        """
        output_csv_path = os.path.join(self.output_dir, output_csv)
        with open(output_csv_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['Segment Type', 'Importance', 'Text'])
            for segment in self.segments:
                writer.writerow([segment.segment_type, segment.importance, segment.text])

    def process(self):
        """
        The main process for segmenting the document and saving the results into a CSV file.

        This function orchestrates the loading, segmenting, and saving of the document.
        """
        try:
            self.segment_document()

            # Generate the CSV file name based on the original file name and current timestamp
            base_filename = os.path.splitext(os.path.basename(self.file_path))[0]
            output_csv = f"{base_filename}_chunk_{datetime.today().strftime('%Y_%m_%d_%H_%M_%S')}.csv"

            # Save the segmented content to the CSV file
            self.save_segments_to_csv(output_csv)
            print(f"Segments have been saved in {os.path.join(self.output_dir, output_csv)}")
        except ValueError as e:
            print(f"Processing error: {e}")

# Utilisation de la classe
if __name__ == "__main__":
    file_path = r'data/original/your document.pdf'  # Remplace par le chemin réel vers ton fichier
    output_dir = 'data/chunk'  # Chemin vers le dossier où sauvegarder le CSV
    segmenter = DocumentSegmenter(file_path, output_dir)
    segmenter.process()
