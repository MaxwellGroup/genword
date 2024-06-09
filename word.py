import os
import sys
import json
from docxtpl import DocxTemplate, RichText

def main(template_path, output_prefix):
    # Verifica l'esistenza del file di template
    if not os.path.isfile(template_path):
        print(f"Errore: Il file di template '{template_path}' non esiste.")
        return

    # Crea un oggetto DocxTemplate dal file di template
    doc = DocxTemplate(template_path)
    print(f"Caricato template: {template_path}")

    # Cerca tutti i file JSON nella directory corrente
    json_files = [f for f in os.listdir('.') if f.endswith('.json')]
    if not json_files:
        print("Nessun file JSON trovato nella directory corrente.")
        return

    print(f"Trovati {len(json_files)} file JSON.")

    # Elabora ogni file JSON
    for i, json_file in enumerate(json_files, start=1):
        json_path = os.path.join('.', json_file)
        if not os.path.isfile(json_path):
            print(f"Errore: Il file JSON '{json_file}' non esiste.")
            continue

        with open(json_path, 'r') as file:
            context = json.load(file)

        # Converti tutte le chiavi del JSON in minuscolo e crea oggetti RichText
        new_context = {k.lower(): RichText(str(v), style='UpperCase') if k.isupper() else RichText(str(v)) for k, v in context.items()}

        # Ottieni i placeholder dal template
        placeholders = doc.get_undeclared_template_variables()

        # Verifica che tutte le chiavi del JSON siano presenti nel template
        missing_keys = [key for key in placeholders if key not in new_context]
        if missing_keys:
            print(f"Errore: Le seguenti chiavi mancano nel file JSON '{json_file}':")
            for key in missing_keys:
                print(f"- {key}")
            continue

        # Esegui il rendering del template con i dati del contesto
        doc.render(new_context)
        print(f"Elaborato file JSON: {json_file}")

        # Salva il documento generato con un nome sequenziale
        output_path_docx = f"{output_prefix}{i:02d}.docx"
        doc.save(output_path_docx)
        print(f"Documento Word generato: {output_path_docx}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Utilizzo: python3 script.py template.docx doc_output")
        sys.exit(1)

    template_path = sys.argv[1]
    output_prefix = sys.argv[2]

    main(template_path, output_prefix)