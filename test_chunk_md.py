import os
import csv
import fitz  # PyMuPDF
from markdownify import markdownify as md
from datetime import datetime
import pymupdf4llm as pymu

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
        """Charge le document PDF et le convertit en texte Markdown."""
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"Le fichier {self.file_path} n'existe pas.")

        if self.file_extension == '.pdf':
            return self._load_pdf_as_markdown()
        else:
            raise ValueError(f"Format de fichier non supporté : {self.file_extension}")

    def _load_pdf_as_markdown(self):
        """Utilise PyMuPDF pour extraire le texte d'un PDF et le convertir en Markdown."""
        markdown_text = pymu.to_markdown(self.file_path)
        # doc = fitz.open(self.file_path)
        # markdown_text = ""
        #
        # # Parcourir toutes les pages du PDF
        # for page_num in range(doc.page_count):
        #     page = doc.load_page(page_num)
        #     text = ("text")  # Récupérer le texte brut
        #     # text = page.get_text("text")  # Récupérer le texte brut
        #     # markdown_text += md(text)  # Convertir en Markdown
        # print(markdown_text)
        return markdown_text

    def segment_document(self):
        """Segmente le document Markdown en titres, sous-titres et paragraphes."""
        markdown_text = self.load_document()

        # Analyse du texte Markdown ligne par ligne
        current_text = ""
        current_segment_type = "paragraph"
        current_importance = 1.0

        for line in markdown_text.splitlines():
            line = line.strip()

            # Détection des titres et sous-titres en Markdown
            if line.startswith('#'):
                # Si un paragraphe précédent est détecté, l'ajouter aux segments
                if current_text and current_segment_type == 'paragraph':
                    self.segments.append(Segment(current_text, current_segment_type, current_importance))

                # Calculer le niveau du titre en fonction du nombre de '#'
                level = len(line.split(' ')[0])

                # Détecter si c'est un titre principal ou un sous-titre
                current_segment_type = 'title' if level == 1 else 'subtitle'
                current_importance = 2.0 if level == 1 else (1.8 if level == 2 else 1.6)

                # Ajouter le titre ou sous-titre au segment
                self.segments.append(Segment(line, current_segment_type, current_importance))

                # Réinitialiser pour le prochain paragraphe
                current_text = ""
                current_segment_type = "paragraph"
                current_importance = 1.0
            else:
                # Si ce n'est pas un titre, c'est un paragraphe
                if line:
                    current_text += line + " "

        # Ajouter le dernier paragraphe s'il y a du texte restant
        if current_text and current_segment_type == 'paragraph':
            self.segments.append(Segment(current_text.strip(), 'paragraph', 1.0))

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
        try:
            self.segment_document()

            # Extraire le nom de base du fichier sans extension
            base_filename = os.path.splitext(os.path.basename(self.file_path))[0]

            # Générer le nom de fichier CSV en utilisant le nom du fichier source
            output_csv = f"{base_filename}_chunk_{datetime.today().strftime('%Y_%m_%d_%H_%M_%S')}.csv"

            self.save_segments_to_csv(output_csv)
            print(f"Les segments ont été sauvegardés dans {os.path.join(self.output_dir, output_csv)}")
        except ValueError as e:
            print(e)


# Utilisation de la classe
if __name__ == "__main__":
    file_path = r'C:\Users\Roc-Antony.COCO\PycharmProjects\PropIAConcept\resources\test_extract_nomme.pdf'  # Chemin vers le fichier PDF de test
    output_dir = 'data/chunk'  # Chemin vers le dossier où sauvegarder le CSV
    segmenter = DocumentSegmenter(file_path, output_dir)
    segmenter.process()
