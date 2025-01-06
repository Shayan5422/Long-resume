# main.py

import os
import subprocess
import logging
import PyPDF2
import pdfplumber
from PIL import Image
from pdf2image import convert_from_path
import pytesseract
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar, LTFigure

# Configuration du logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def run_Qwen32_avec_texte(prompt, texte):
    """
    Envoie le texte au modèle LLM Qwen2.5:32b et retourne le résumé. 
    """
    command = ['ollama', 'run', 'llama3.2:3b']
    full_prompt = f"{prompt}\n\n{texte}"
    
    try:
        process = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        output, error = process.communicate(input=full_prompt)
        
        if process.returncode != 0:
            logger.error(f"Erreur lors de l'exécution du modèle : {error}")
            return None
        
        return output.strip()
    except Exception as e:
        logger.error(f"Exception lors de l'appel au modèle LLM : {e}")
        return None

def extraire_texte_du_pdf_seulement(chemin_pdf):
    """
    Extrait le texte d'un fichier PDF et retourne un dictionnaire avec le texte par page.
    """
    texte_par_page = {}
    drapeau_image = False
    try:
        pdf_file = open(chemin_pdf, 'rb')
        pdf_reader = PyPDF2.PdfReader(pdf_file)
    except Exception as e:
        logger.error(f"Échec de l'ouverture du PDF {chemin_pdf} : {e}")
        return {}
    
    try:
        for num_page, layout_page in enumerate(extract_pages(chemin_pdf)):
            try:
                page_obj = pdf_reader.pages[num_page]
            except IndexError:
                logger.warning(f"La page {num_page} n'existe pas dans le PDF {chemin_pdf}.")
                continue
            
            texte_page = []
            texte_images = []
            contenu_page = []
            
            try:
                pdf = pdfplumber.open(chemin_pdf)
                tables_page = pdf.pages[num_page]
                tables = tables_page.find_tables()
            except Exception as e:
                logger.error(f"Échec de l'extraction des tables de la page {num_page} : {e}")
                tables = []
            
            # Extraction des tables
            for num_table in range(len(tables)):
                try:
                    table = extraire_table(chemin_pdf, num_page, num_table)
                    table_string = convertir_table(table)
                    contenu_page.append(table_string)
                except Exception as e:
                    logger.error(f"Échec de l'extraction de la table {num_table} de la page {num_page} : {e}")
                    continue
            
            # Tri des éléments par position Y (descendant)
            elements_page = [(element.y1, element) for element in layout_page]
            elements_page.sort(key=lambda a: a[0], reverse=True)
            
            for composant in elements_page:
                element = composant[1]
                
                # Extraction des éléments de texte
                if isinstance(element, LTTextContainer):
                    try:
                        ligne_texte, _ = extraction_texte(element)
                        texte_page.append(ligne_texte)
                        contenu_page.append(ligne_texte)
                    except Exception as e:
                        logger.error(f"Échec de l'extraction du texte de l'élément sur la page {num_page} : {e}")
                
                # Extraction des éléments d'image (si présents)
                if isinstance(element, LTFigure):
                    try:
                        rogner_image(element, page_obj)
                        convertir_en_image('cropped_image.pdf')
                        texte_image = image_vers_texte('PDF_image.png')
                        texte_images.append(texte_image)
                        contenu_page.append(texte_image)
                        drapeau_image = True
                    except Exception as e:
                        logger.error(f"Échec de l'extraction du texte de l'image sur la page {num_page} : {e}")
            
            # Combiner le contenu extrait pour chaque page
            clef = f'Page_{num_page + 1}'  # Pages commencent à 1 pour l'utilisateur
            texte_par_page[clef] = contenu_page
    
    except Exception as e:
        logger.error(f"Échec de l'extraction du texte du PDF {chemin_pdf} : {e}")
    finally:
        pdf_file.close()
        if drapeau_image:
            if os.path.exists('cropped_image.pdf'):
                os.remove('cropped_image.pdf')
            if os.path.exists('PDF_image.png'):
                os.remove('PDF_image.png')
    
    return texte_par_page

def extraire_table(chemin_pdf, num_page, num_table):
    """
    Extrait une table spécifique d'une page PDF.
    """
    try:
        pdf = pdfplumber.open(chemin_pdf)
        table_page = pdf.pages[num_page]
        table = table_page.extract_tables()[num_table]
        return table
    except Exception as e:
        logger.error(f"Échec de l'extraction de la table {num_table} de la page {num_page + 1} : {e}")
        return []

def convertir_table(table):
    """
    Convertit une table en chaîne de caractères formatée.
    """
    table_string = ''
    try:
        for ligne in table:
            ligne_nettoyee = [item.replace('\n', ' ') if item is not None and '\n' in item else ('None' if item is None else item) for item in ligne]
            table_string += ('|' + '|'.join(ligne_nettoyee) + '|\n')
        table_string = table_string.rstrip('\n')
        return table_string
    except Exception as e:
        logger.error(f"Échec de la conversion de la table en chaîne de caractères : {e}")
        return ""

def rogner_image(element, page_obj):
    """
    Rogne une image à partir des coordonnées de l'élément.
    """
    image_gauche, image_haut, image_droite, image_bas = element.x0, element.y0, element.x1, element.y1
    page_obj.mediabox.lower_left = (image_gauche, image_bas)
    page_obj.mediabox.upper_right = (image_droite, image_haut)
    
    writer_cropped = PyPDF2.PdfWriter()
    writer_cropped.add_page(page_obj)
    with open('cropped_image.pdf', 'wb') as fichier_cropped:
        writer_cropped.write(fichier_cropped)

def convertir_en_image(fichier_entree):
    """
    Convertit un fichier PDF en image PNG.
    """
    try:
        images = convert_from_path(fichier_entree)
        if images:
            image = images[0]
            fichier_sortie = 'PDF_image.png'
            image.save(fichier_sortie, 'PNG')
            logger.info(f"Converti {fichier_entree} en image {fichier_sortie}.")
    except Exception as e:
        logger.error(f"Échec de la conversion de {fichier_entree} en image : {e}")

def image_vers_texte(chemin_image):
    """
    Utilise OCR pour extraire le texte d'une image.
    """
    try:
        img = Image.open(chemin_image)
        texte = pytesseract.image_to_string(img)
        logger.info(f"Texte extrait de l'image {chemin_image}.")
        return texte
    except Exception as e:
        logger.error(f"Échec de l'extraction du texte de l'image {chemin_image} : {e}")
        return ""

def extraction_texte(element):
    """
    Extrait le texte et les formats des caractères d'un élément de texte.
    """
    ligne_texte = element.get_text()
    formats_ligne = []
    for text_line in element:
        if isinstance(text_line, LTChar):
            formats_ligne.append(text_line.fontname)
            formats_ligne.append(text_line.size)
    formats_par_ligne = list(set(formats_ligne))
    return (ligne_texte, formats_par_ligne)

def extraire_texte_du_dossier(chemin_dossier, chemin_sortie_txt):
    """
    Extrait le texte de tous les fichiers PDF dans un dossier et les enregistre dans un fichier texte.
    """
    texte_combine = ''
    for nom_fichier in os.listdir(chemin_dossier):
        if nom_fichier.endswith('.pdf'):
            chemin_pdf = os.path.join(chemin_dossier, nom_fichier)
            logger.info(f"Extraction du texte de : {chemin_pdf}")
            texte_pdf = extraire_texte_du_pdf_seulement(chemin_pdf)
            texte_combine += f"\n--- Texte de {nom_fichier} ---\n"
            for clef, contenu in texte_pdf.items():
                texte_combine += f"\n*** {clef} ***\n" + '\n'.join(contenu) + "\n"
    
    # Enregistrer le texte combiné dans un fichier .txt
    try:
        with open(chemin_sortie_txt, 'w', encoding='utf-8') as fichier_sortie:
            fichier_sortie.write(texte_combine)
        logger.info(f"Tout le texte extrait a été sauvegardé dans {chemin_sortie_txt}")
    except Exception as e:
        logger.error(f"Échec de l'écriture du texte extrait dans {chemin_sortie_txt} : {e}")
    
    return texte_combine

def main():
    """
    Fonction principale qui orchestre l'extraction, le résumé et la sauvegarde des résultats.
    """
    chemin_pdf = "/Users/shayanhashemi/Downloads/extract_long/Partenariats exemple/RA Arche SMR 2023.pdf"  # Remplacez par le chemin de votre PDF
    chemin_sortie = "résumés.txt"  # Fichier de sortie pour les résumés
    prompt = "Veuillez fournir un résumé concis de la section suivante :"
    
    # Extraire le texte par page
    logger.info(f"Début de l'extraction du texte de {chemin_pdf}")
    texte_par_page = extraire_texte_du_pdf_seulement(chemin_pdf)
    
    if not texte_par_page:
        logger.error("Aucun texte extrait. Fin du programme.")
        return
    
    résumés = {}
    
    # Résumer chaque page
    for clef, contenu in texte_par_page.items():
        texte_page = '\n'.join(contenu)
        logger.info(f"Envoi de {clef} au modèle LLM pour résumé.")
        résumé = run_Qwen32_avec_texte(prompt, texte_page)
        if résumé:
            résumés[clef] = résumé
            logger.info(f"Résumé de {clef} obtenu.")
        else:
            résumés[clef] = "Résumé non disponible en raison d'une erreur."
            logger.warning(f"Résumé de {clef} non disponible.")
    
    # Enregistrer les résumés dans un fichier
    try:
        with open(chemin_sortie, 'w', encoding='utf-8') as fichier_sortie:
            for clef, résumé in résumés.items():
                fichier_sortie.write(f"{clef}:\n{résumé}\n\n")
        logger.info(f"Tous les résumés ont été sauvegardés dans {chemin_sortie}")
    except Exception as e:
        logger.error(f"Échec de l'écriture des résumés dans {chemin_sortie} : {e}")

if __name__ == "__main__":
    main()
