import morse
import time
from docx import Document

mensaje = input("Dime algo: ")

telegrama = morse.toMorse(mensaje)
print(telegrama)
original = morse.toPlain(telegrama)
print(original)

print(time.strftime("%d/%m/%Y", time.gmtime()))