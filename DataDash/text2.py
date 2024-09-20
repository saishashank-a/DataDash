import pandas as pd
import difflib
import os

def get_word_diff_details(text1, text2):
    """Compare two texts and return the word-level differences"""
    seq_matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    
    changes = []
    change_positions = []
    
    for opcode in seq_matcher.get_opcodes():
        if opcode[0] != 'equal':
            tag, i1, i2, j1, j2 = opcode
            if tag == 'replace':
                changes.append(f"'{' '.join(text1.split()[i1:i2])}' replaced with '{' '.join(text2.split()[j1:j2])}'")
            elif tag == 'delete':
                changes.append(f"'{' '.join(text1.split()[i1:i2])}' deleted")
            elif tag == 'insert':
                changes.append(f"'{' '.join(text2.split()[j1:j2])}' inserted")
            change_positions.append(f"{i1}-{i2}")

    return changes, len(changes), change_positions

# Path to your Excel file
file_path = 'worddifference.xlsx'

# Check if the file exists
if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
else:
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Initialize new columns
    df['Word_Changes'] = ''
    df['Number_of_Word_Changes'] = 0
    df['Change_Positions'] = ''

    # Iterate and compare the text in Month 1 and Month 2
    for i, row in df.iterrows():
        text1 = str(row['M1_Text Content']) if pd.notna(row['M1_Text Content']) else ''
        text2 = str(row['M2_Text Content']) if pd.notna(row['M2_Text Content']) else ''
        
        changes, num_changes, change_positions = get_word_diff_details(text1, text2)

        # Update the dataframe with the changes
        df.at[i, 'Word_Changes'] = ' | '.join(changes)
        df.at[i, 'Number_of_Word_Changes'] = num_changes
        df.at[i, 'Change_Positions'] = ', '.join(change_positions)

    # Save updated dataframe back to Excel
    df.to_excel('updated_word_changes.xlsx', index=False)
    print("File processed and saved as 'updated_word_changes.xlsx'")