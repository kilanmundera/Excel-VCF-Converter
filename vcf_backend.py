import pandas as pd

# This is the shared configuration dictionary used to pass data from the GUI
CONFIG = {}

def vcf_escape(text):
    """
    Escapes characters in a string as required by the VCF standard (RFC 2426/6350).
    Escapes backslashes, semicolons, and commas.
    """
    if not isinstance(text, str):
        text = str(text)
        
    # Order matters: Backslash must be escaped first.
    text = text.replace('\\', r'\\')
    text = text.replace(';', r'\;')
    text = text.replace(',', r'\,')
    
    return text

def create_vcf_from_excel(config):
    """
    Creates VCF from an in-memory DataFrame, a column mapping,
    and a custom Full Name (FN) template.
    """
    
    # 1. Retrieve data and configuration
    try:
        df = config["dataframe_to_process"]
        mapping = config["column_mapping"]
        output_file = config["output_vcf_filename"]
        fn_template = config["fn_template"]
    except KeyError as e:
        raise Exception(f"Configuration is missing a required key: {e}. Cannot proceed with conversion.")

    if not mapping:
        raise Exception("Column mapping is empty. Please configure it in the app.")
    if not fn_template:
        raise Exception("Name Structure template is empty.")

    vcf_cards = []
    
    # 2. Iterate over the DataFrame rows
    for _, row in df.iterrows():
        # Initialize VCF card
        card = ["BEGIN:VCARD", "VERSION:3.0"]
        
        # Get all mapped values from the row, applying VCF escaping immediately.
        vals = {
            "Full Name": vcf_escape(row.get(mapping.get("Full Name"), "")),
            "Title":     vcf_escape(row.get(mapping.get("Title"), "")),
            "Phone":     vcf_escape(row.get(mapping.get("Phone"), "")),
            "Email":     vcf_escape(row.get(mapping.get("Email"), "")),
            # Notes is escaped here, but newlines must be handled specifically later.
            "Notes":     vcf_escape(row.get(mapping.get("Notes"), "")), 
        }

        # --- Generate FN (Formatted Name) ---
        final_fn = fn_template
        
        # Replace placeholders with the escaped values
        for field_name, value in vals.items():
            placeholder = f"{{{{{field_name}}}}}"
            final_fn = final_fn.replace(placeholder, value)
            
        # Clean up excessive whitespace often left by empty fields
        final_fn = final_fn.replace("  ", " ").strip()
        
        # If the string is empty or only contains separators, clear it.
        if all(c in " -/|()[]{}" for c in final_fn):
             final_fn = ""

        # --- Add FN and N properties ---
        if final_fn:
            card.append(f"FN:{final_fn}")
            
            # Structured Name (N): Fallback to FN and set other components to empty
            card.append(f"N:{final_fn};;;;")
        
        # --- Add other fields ---
        if vals["Title"]:
            card.append(f"TITLE:{vals['Title']}")
            
        if vals["Phone"]:
            card.append(f"TEL;TYPE=CELL:{vals['Phone']}")
            
        if vals["Email"]:
            card.append(f"EMAIL;TYPE=INTERNET:{vals['Email']}")
            
        if mapping.get("Notes") and row.get(mapping.get("Notes")):
            # Retrieve the raw note content again to handle newlines
            raw_note = str(row.get(mapping["Notes"]))
            
            # Newlines are specifically encoded as '\n' in the NOTE property
            note_with_newlines_encoded = raw_note.replace("\n", "\\n")
            
            # Apply VCF escaping to the result
            final_note = vcf_escape(note_with_newlines_encoded)
            card.append(f"NOTE:{final_note}")
        
        # VCF End
        card.append("END:VCARD")
        
        # Only add cards that contain actual contact data (more than BEGIN/VERSION/END)
        if len(card) > 3: 
            vcf_cards.append("\n".join(card))

    if not vcf_cards:
        raise Exception(f"No valid contacts were generated from {len(df)} rows. Check your column mapping and Excel data.")

    # 3. Write all cards to the output file
    # Ensures UTF-8 encoding and proper card separation with trailing newlines.
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("\n\n".join(vcf_cards) + "\n\n")
