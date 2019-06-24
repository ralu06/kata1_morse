import morse
from tkinter import *
from tkinter import ttk
from tkinter import messagebox

class Translator(ttk.Frame):
    def __init__(self, parent, **kwargs):
        ttk.Frame.__init__(self, parent, height=kwargs['height'], width=kwargs['width'])

        self.traduccionDirecta = True

        self.sender = StringVar()
        self.receiver = StringVar()

        sender_lbl = ttk.Label(self,text="Sender:", width=11, anchor=W)
        sender_lbl.place(x=12, y=12)
        self.sender_entry = ttk.Entry(self, width=16, textvariable=self.sender)
        self.sender_entry.place(x=70, y=10)

        receiver_lbl = ttk.Label(self,text="Receiver:", width=11, anchor=W)
        receiver_lbl.place(x=250, y=12)
        self.receiver_entry = ttk.Entry(self, width=16, textvariable=self.receiver)
        self.receiver_entry.place(x=318, y=10)

        self.origin_lbl = ttk.Label(self, text="Plano", width=6, anchor=W)
        self.origin_lbl.place(x=12, y=40)
        self.origin_text = Text(self, width=26, height=8)
        self.origin_text.place(x=32, y=60)

        self.origin_text.bind("<KeyRelease>", self.traduce)

        self.destino_lbl = ttk.Label(self, text="Morse")
        self.destino_lbl.place(x=250, y=40)
        self.destino_text = Text(self, width=27, height=8)
        self.destino_text.place(x=270, y=60)

        btn_send = ttk.Button(self,command=self.send_telegram, text="Send")
        btn_send.place(y=165, x=500)

        btn_change = ttk.Button(self, command=self.changeText, text="<=>")
        btn_change.place(y=140, x=500)

    def traduce(self, e):
        
        texto_original = self.origin_text.get("1.0","end-1c")
        if self.traduccionDirecta:
            traduccion = morse.toMorse(texto_original)
        else:
            traduccion = morse.toPlain(texto_original)
        self.destino_text.delete("1.0", END)
        self.destino_text.insert(INSERT, traduccion)

        print(traduccion)   

    def send_telegram(self):
        sender = self.sender.get()
        receiver = self.receiver.get()

        if self.traduccionDirecta:
            msg = self.origin_text.get("1.0", "end-1c")
        else:
            msg = self.destino_text.get("1.0", "end-1c")

        if sender.strip == "":    
            messagebox.showwarning()
            messagebox.showwarning ("Aviso", "Debe informar el sender", parent = self)
            return

        if receiver.strip == "":
            messagebox.showwarning ("Aviso", "Debe informar el receiver", parent = self)
            return

        if msg.strip == "":
            messagebox.showwarning ("Aviso", "Debe informar el mensaje", parent = self)
            return
       
    
    def changeText(self):
        if self.traduccionDirecta:
            self.origin_lbl.config(text='Morse')
            self.destino_lbl.config(text='Plano')
        else:
            self.origin_lbl.config(text='Plano')
            self.destino_lbl.config(text='Morse')

        self.traduccionDirecta = not self.traduccionDirecta
        

class MainApp(Tk):
    def __init__(self):
        Tk.__init__(self)
        self.title("Traductor Morse")
        self.geometry("600x200")

        self.translator = Translator(self, height=200, width=600)
        self.translator.place(x=0, y=0)

    def start(self):
        self.mainloop()

if __name__ == "__main__":
    miTraductor = MainApp()
miTraductor.start()