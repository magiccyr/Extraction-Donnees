# 📄 Convertisseur Universel PDF ⇄ Word ⇄ Excel
## 📌 Présentation

Ce projet est une application Python dotée d’une interface graphique (Tkinter) permettant de :

- Extraire des données d’un PDF vers Excel

- Convertir un Excel vers PDF

- Convertir un PDF vers Word (.docx)

- Convertir un Word vers PDF

- Afficher un aperçu avant chaque conversion

- Gérer automatiquement les tableaux, le texte brut et les paires clé/valeur

- Ce convertisseur universel facilite la manipulation de documents dans différents formats professionnels.

## 🚀 Fonctionnalités
### PDF → Excel

- Extraction automatique :

  - tableaux (via pdfplumber)

  - texte brut
  - couples clé : valeur

- Affichage d’un tableau preview

- Export en .xlsx

### Excel → PDF

- Chargement et prévisualisation d'un classeur Excel

- Conversion en fichier PDF formaté via ReportLab

- Support jusqu'à 100 lignes affichées dans le PDF

### PDF → Word

- Extraction du texte page par page (PyPDF2)

- Options :

  - conservation du formatage

  - ajout automatique de sauts de page

- Génération d’un .docx structuré

### Word → PDF

- Lecture du document avec python-docx

- Conversion en PDF via ReportLab

- Prise en charge des titres, paragraphes, mise en forme simple

## 🛠️ Installation
### 1. Prérequis

Assurez-vous d’avoir Python 3.9+ installé.

### 2. Installer les dépendances
```bash
pip install pdfplumber PyPDF2 reportlab python-docx openpyxl pandas
```

Si certaines dépendances manquent, l’application affichera automatiquement un avertissement.

## 📁 Architecture du code
Le fichier principal contient :
### ✔️ UniversalConverterApp

Classe principale qui :

- initialise l’interface

- gère les onglets et widgets

- appelle les fonctions de conversion

### ✔️ Fonctions principales

- <b>extract_pdf_data()</b> : extraction PDF → DataFrame

- <b>export_to_excel()</b> : export vers Excel

- <b>convert_excel_to_pdf()</b> : mise en page PDF (Reportlab)

- <b>convert_pdf_to_word()</b> : conversion PDF → Word

- <b>convert_word_to_pdf()</b> : conversion Word → PDF

- <b>display_dataframe()</b> : affichage des DataFrames dans un TreeView

### ✔️ Compatibilité étendue

- <b>pdfplumber</b> pour extraction structurée

- <b>PyPDF2</b> pour lecture des pages

- <b>reportlab</b> pour création PDF

- <b>python-docx</b> pour Word

## ⚠️ Limitations connues

- Les PDF scannés ou images non OCR ne sont pas analysés (pas de reconnaissance de caractères).

- Le formatage complexe (tableaux Word, images, styles avancés) peut ne pas être parfaitement reproduit.

- Certaines polices ou langues spéciales peuvent nécessiter des fonts adaptées côté reportlab.

## 🧩 Améliorations possibles

- ✔️ Ajouter un OCR (Tesseract) pour les PDF scannés
- ✔️ Ajouter la conversion Word → Excel ou PDF → CSV
- ✔️ Export multi-feuilles pour Excel
- ✔️ Interface modernisée (customtkinter)

## 👤 Auteur

Projet développé par <b>Cyr DJOKI</b> pour faciliter la conversion multi-format avec une interface simple, efficace et extensible.
