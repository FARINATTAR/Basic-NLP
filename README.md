# **Extract POS from DOCX**

## **Description**
**This Python script extracts and categorizes parts of speech (POS) from a Word document (.docx).** The script reads the content of the input document, tokenizes the text, removes stopwords, and then categorizes words into nouns, verbs, and adjectives. The extracted words are saved into a new Word document for easy viewing and further analysis.

## **Features**
- **Reads text from a .docx Word document.**
- **Tokenizes the text and removes common stopwords.**
- **Identifies parts of speech (POS) for each word (Nouns, Verbs, Adjectives).**
- **Saves the categorized results (Nouns, Verbs, Adjectives) into a new .docx file.**

## **Requirements**
Before using the script, you need to install the following libraries:

- **nltk** for natural language processing tasks like tokenization and POS tagging.
- **python-docx** for handling .docx files.

You can install these dependencies using pip:

```bash
pip install nltk python-docx

NLTK Setup
The script requires certain NLTK resources for tokenization and stopwords. You can download them by running the following commands within a Python environment:
import nltk
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('stopwords')


### Key Points:
- **Only modify** the `input_file` and `output_file` paths in the script.
- **Run the script** to process the document and save the extracted POS categories into a new Word document.

This way, the user will clearly understand that only the file paths need to be adjusted.
