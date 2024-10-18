import os
import csv
import textract  # Pour extraire le texte des PDF et d'autres formats
from docx import Document
from datetime import datetime
import fitz  # PyMuPDF
import pdfplumber

class Segment:
    def __init__(self, text, segment_type, importance=1.0):
        self.text = text
        self.segment_type = segment_type
        self.importance = importance

    def __repr__(self):
        return f"Segment(type={self.segment_type}, importance={self.importance}, text={self.text[:30]}...)"

class DocumentSegmenter:
    def __init__(self, file_path, output_dir):
        self.file_path = file_path
        self.segments = []
        self.file_extension = os.path.splitext(file_path)[1].lower()
        self.output_dir = output_dir

        # Vérifier si le dossier de sortie existe, sinon le créer
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            print(f"Dossier de sauvegarde créé : {self.output_dir}")
        else:
            print(f"Dossier de sauvegarde existant : {self.output_dir}")

    def load_document(self):
        """Charge le document en fonction de son type (DOCX, PDF, TXT)."""
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"Le fichier {self.file_path} n'existe pas.")

        if self.file_extension == '.docx':
            return self._load_docx()
        elif self.file_extension == '.pdf':
            return self._load_pdf()
        elif self.file_extension == '.txt':
            return self._load_txt()
        else:
            raise ValueError(f"Format de fichier non supporté : {self.file_extension}")

    def _load_docx(self):
        """Charge un document DOCX et retourne son contenu en tant que texte brut."""
        doc = Document(self.file_path)
        text = []
        for paragraph in doc.paragraphs:
            text.append(paragraph.text)
        return "\n".join(text)

    def _load_pdf(self):
        """Utilise pdfplumber pour extraire le texte d'un fichier PDF en fonction de la taille de la police."""
        text_segments = []
        try:
            with pdfplumber.open(self.file_path) as pdf:
                for page in pdf.pages:
                    for element in page.extract_words():
                        font_size = float(element['height'])
                        if font_size >= 16:
                            # Titre
                            text_segments.append(Segment(element['text'], 'title', 2.0))
                        elif 12 <= font_size < 16:
                            # Sous-titre
                            text_segments.append(Segment(element['text'], 'subtitle', 1.8))
                        else:
                            # Paragraphe
                            text_segments.append(Segment(element['text'], 'paragraph', 1.0))
            return text_segments if text_segments else None
        except Exception as e:
            print(f"Erreur lors de la lecture du fichier PDF : {e}")
            return None

    def _load_txt(self):
        """Lit un fichier TXT et retourne son contenu en tant que texte brut."""
        with open(self.file_path, 'r', encoding='utf-8') as file:
            return file.read()

    def segment_document(self):
        """Segmente le document chargé en fonction des styles et du type de texte."""
        text = self.load_document()

        if self.file_extension == '.docx':
            self._segment_docx(text)
        else:
            self._segment_plain_text(text)

        # Add remaining text as a paragraph segment
        if text.strip():
            self.segments.append(Segment(text.strip(), "paragraph", 1.0))


    def _segment_docx(self, text):
        """Segmente un fichier DOCX avec les styles de titre, sous-titre et paragraphe."""
        doc = Document(self.file_path)
        current_text = ""
        current_segment_type = "paragraph"
        current_importance = 1.0

        for para in doc.paragraphs:
            line = para.text.replace('\u00A0', ' ').strip()

            if not line:
                continue  # Ignorer les paragraphes vides

            if para.style.name.startswith('Heading'):
                if current_text and current_segment_type == 'paragraph':
                    self.segments.append(Segment(current_text, current_segment_type, current_importance))

                level = int(para.style.name.split()[1])
                current_segment_type = 'title' if level == 1 else 'subtitle'
                current_importance = 1.0 if level == 1 else 0.8
                self.segments.append(Segment(line, current_segment_type, current_importance))

                current_text = ""
                current_segment_type = "paragraph"
                current_importance = 1.0
            else:
                current_text += line + " "

        if current_text:
            self.segments.append(Segment(current_text.strip(), current_segment_type, current_importance))

    def _segment_plain_text(self, text):
        """Segmente un texte brut (PDF, TXT) en paragraphes."""
        paragraphs = text.split("\n")
        for paragraph in paragraphs:
            if paragraph.strip():
                self.segments.append(Segment(paragraph.strip(), "paragraph", 1.0))

    def save_segments_to_csv(self, output_csv):
        """Sauvegarde les segments dans un fichier CSV."""
        output_csv_path = os.path.join(self.output_dir, output_csv)
        with open(output_csv_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['Segment Type', 'Importance', 'Text'])
            for segment in self.segments:
                writer.writerow([segment.segment_type, segment.importance, segment.text])

    def process(self):
        """Processus complet de segmentation et de sauvegarde dans un fichier CSV."""
        self.segment_document()

        # Extraire le nom de base du fichier sans extension
        base_filename = os.path.splitext(os.path.basename(self.file_path))[0]

        # Générer le nom de fichier CSV en utilisant le nom du fichier source
        output_csv = f"{base_filename}_chunk_{datetime.today().strftime('%Y_%m_%d_%H_%M_%S')}.csv"

        self.save_segments_to_csv(output_csv)

# Utilisation de la classe
if __name__ == "__main__":
    # file_path = r'C:\Users\Roc-Antony.COCO\PycharmProjects\PropIAConcept\resources\consultation_TMA_Feg_2024_Lot_3_V1.0.docx'  # Remplace par le chemin réel vers ton fichier (DOCX, PDF, TXT)
    file_path = r'C:\Users\Roc-Antony.COCO\PycharmProjects\PropIAConcept\resources\test_extract_nomme.pdf'  # Remplace par le chemin réel vers ton fichier (DOCX, PDF, TXT)
    output_dir = 'data/chunk'  # Chemin vers le dossier où sauvegarder le CSV
    segmenter = DocumentSegmenter(file_path, output_dir)
    segmenter.process()
