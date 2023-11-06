from docx import Document
import re
import os
import openpyxl
from datetime import datetime



transcripts_folder_path = os.getcwd()+'/raw_transcripts/'
# box location on my mac /Users/mahdijafari/Library/CloudStorage/Box-Box


# # Define a function to extract the time from the string
# def extract_time_from_string(time_string):
#     # Split the string by spaces
#     parts = time_string.split()
#     for part in parts:
#         try:
#             # Try to parse a time from the part
#             time = datetime.strptime(part, "%H:%M:%S")
#             return time.time()
#         except ValueError:
#             print(ValueError)
#             pass
#     return None

def find_doc_files_in_current_folder() -> list:
    """Find the docx file in the current folder

    Returns:
        list: list of the files name as str
    """

    doc_files = []

    for filename in os.listdir(transcripts_folder_path):
        if filename.endswith(".docx"):
            doc_files.append(filename)

    return doc_files

list_of_docx_files_name = find_doc_files_in_current_folder()


def find_speaker_change(file_name : str) -> list:

    # Create a list to store discussions
    discussions = []
    current_discussion = ""
    
    doc = Document(transcripts_folder_path+file_name)
    # Define a regular expression pattern to match speaker lines
    speaker_pattern = re.compile(r'\d{2}:\d{2}:\d{2} Speaker \d')

    for paragraph in doc.paragraphs:
        text = paragraph.text

        # Check if the paragraph matches the speaker pattern
        if re.match(speaker_pattern, text):
            # Start a new discussion when a new speaker is detected
            if current_discussion:
                discussions.append(current_discussion)
                
            current_discussion = text + "\n"
        else:
            # Append the text to the current discussion
            current_discussion += text + "\n"


    # Append the last discussion to the list
    if current_discussion:
        discussions.append(current_discussion)

    pattern = r'\bSpeaker \d+\b'
    temp_speaker_numbers = ''
    highest_speaker = 0
    discussion_number = 0
    speaker_time_dict = {}

    for discussion in discussions:

        # extracting speakers from the file, and ignore same speaker in the discussion
        if re.match(speaker_pattern, discussion):
            speaker_numbers = re.findall(pattern, discussion)
            
            if temp_speaker_numbers == '':
                discussion_number +=1
                temp_speaker_numbers = speaker_numbers
                # print("new speaker detected!")
                # print(speaker_numbers)
            elif speaker_numbers != temp_speaker_numbers:
                discussion_number +=1
                temp_speaker_numbers = speaker_numbers
                # print("new speaker detected!")
                # print(speaker_numbers)

            #finding the number of the speakers in the meeting
            num_pattern = r'\d+'
            spk_numbers = int(re.findall(num_pattern, speaker_numbers[0])[0])
            if highest_speaker < spk_numbers:
                highest_speaker = spk_numbers
                
            
            # # buidling disctionary for speaker time per session
            # time_format = "%H:%M:%S"
            # if speaker_numbers[0] not in speaker_time_dict:
            #     speaker_with_time = re.findall(r'\d{2}:\d{2}:\d{2} Speaker \d', discussion)[0]

            #     time_1 = extract_time_from_string(speaker_with_time)
            #     print(time_1)
            #     speaker_time_dict[speaker_numbers[0]] = 0
    # print(speaker_time_dict)

    print(f"file: {file_name} has {discussion_number} discussion.")
    

    return [file_name, discussion_number, highest_speaker]



workbook = openpyxl.Workbook()
sheet = workbook.active
excel_row = ['name', 'discussion number', 'highest speaker number']
sheet.append(excel_row)
for file_name in list_of_docx_files_name:
    excel_row = find_speaker_change(file_name)
    sheet.append(excel_row)

workbook.save("ACIP meeting discussion analysis.xlsx")


{
    "speaker 1": [7000, 7],
    "speaker 2": [5000, 3],
    "after covid" : 1,
    "before covid": 0
}