import openpyxl
import math

def calculate_media(student_notes): 
    media = 0
    
    for note in student_notes:
        media += note.value
        
    return (media / 3) / 10   


def get_total_lessons(row_text):
    total_lessons_text = row_text
    total_lessons = int(total_lessons_text[1])
    
    return total_lessons
    

def alter_row(row, situation, required_note):
    row[6].value = situation
    row[7].value = required_note
    
    return row 
    
def insert_situation(row, media, absences_percentage): 
    if absences_percentage > 0.25: 
        row = alter_row(row, "Reprovado por falta", 0)
    else:
        if media >= 7:
            row = alter_row(row, "Aprovado", 0)
        elif media >= 5 and media < 7:
            naf = math.ceil((media + 5) / 2)
            row = alter_row(row, "exame final", naf)
        else:
            row = alter_row(row, "Reprovado", 0)
        
            
def app():
    try:
        workbook = openpyxl.load_workbook("Cópia de Engenharia de Software - Desafio Fabrício Sobrinho.xlsx")
        eng_sheet = workbook["engenharia_de_software"]

        total_lessons = get_total_lessons(eng_sheet["A2"].value.split(": "))
        for row in eng_sheet.iter_rows(min_row=4):    
            student_notes = []
            student_notes.extend([row[3], row[4], row[5]])
                
            media = calculate_media(student_notes)
            
            absences = row[2].value
            absences_percentage = (absences / total_lessons)
            
            insert_situation(row, media,absences_percentage)   
            
        workbook.save("Cópia de Engenharia de Software - Desafio Fabrício Sobrinho.xlsx")
        print("As informações foram inseridas com êxito!")
    except: 
        print("Houve um erro ao inserir as informações na tabela.")
        

app()


