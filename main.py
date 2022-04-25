#Importing tk libraty
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog as fd
# Importing the PIL library
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import fitz
import os
import win32com.client
import pathlib

# root window
root = tk.Tk()
root.geometry("360x450")
root.title('Štambilj')
root.resizable(0, 0)

# configure the grid
root.columnconfigure(0, weight=1)
#root.columnconfigure(1, weight=3)
#root.rowconfigure(0, weight=4)
#root.rowconfigure(3, weight=4)

#prozor za učitavanje word file
def select_word_file():
    info_poruka.set("")
    filetypes = (
        ('word files', '*.doc'),
        ('word files', '*.docx')
    )

    filename_word = fd.askopenfilename(
        title='Učitaj dokument',
        initialdir='/',
        filetypes=filetypes)
    global proba_word
    proba_word = filename_word


# Prozor za učitavanje prijave

def select_file():
    info_poruka.set("")
    filetypes = (
        ('PDF files', '*.pdf'),
    )

    filename = fd.askopenfilename(
        title='Učitaj prijavu',
        initialdir='/',
        filetypes=filetypes)
    global proba_pdf
    proba_pdf = filename

##### CONVERTING WORD TO PDF (first label frame) ######

radna_povrsina = os.path.join((os.environ['USERPROFILE']), 'Desktop')

def convert_word_to_pdf():
    try:
        input_word_file = str(proba_word)
        input_word = input_word_file.replace("/", "\\")
        radna_povrsina = os.path.join((os.environ['USERPROFILE']), 'Desktop')
        from subprocess import Popen
        LIBRE_OFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"

        p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
                   radna_povrsina, input_word])
        print([LIBRE_OFFICE, '--convert-to', 'pdf', input_word])
        p.communicate()
        info_poruka.set("Uspešno konvertovan word dokument u PDF!")
    except:
        info_poruka.set("Nešto nije u redu, pokušajte ponovo.")


##first label frame for konvertor word to pdf
lf1 = ttk.LabelFrame(root, text='Konvertor word u pdf', width=320, height=160)
lf1.columnconfigure(0, weight=1)
lf1.columnconfigure(1, weight=2)
lf1.grid(column=0, row=0, padx=20, pady=20, sticky=(N, S, E, W))
#Button for open word file
open_button_word = ttk.Button(
    lf1,
    text='Učitaj word',
    command=select_word_file
)

open_button_word.grid(column=0, row=0, padx=5, pady=5, sticky=(N, S, E, W))
#Button for convert word file to pdf
convert_button_word = ttk.Button(
    lf1,
    text='Konvertuj word u pdf',
    command=convert_word_to_pdf
)

convert_button_word.grid(column=1, row=0, padx=5, pady=5, sticky=(N, S, E, W))

##second label frame for stavi stambilj
lf2 = ttk.LabelFrame(root, text='Stavi štambilj', width=320, height=160)
lf2.columnconfigure(0, weight=1)
lf2.columnconfigure(1, weight=2)
lf2.grid(column=0, row=1, padx=20, pady=20, sticky=(N, S, E, W))

##third label frame for info messages
lf3 = ttk.LabelFrame(root, text='Info', width=340, height=160)
lf3.grid(column=0, row=2, padx=20, pady=20, sticky=(N, S, E, W))

#label with info massages
info_poruka = StringVar()
rez= ttk.Label(lf3, textvariable=info_poruka)
rez.grid(row=0, column=0, sticky = (W, E), padx = 10, pady = 10)

# Dugme za učitavanje prijave

open_button = ttk.Button(
    lf2,
    text='Učitaj prijavu',
    command=select_file
)

open_button.grid(column=0, row=0, columnspan = 2, sticky=tk.EW, padx=5, pady=5)

# datum prijave
datum_label = ttk.Label(lf2, text="Datum prijave")
datum_label.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)

datum_entry = ttk.Entry(lf2)
datum_entry.grid(column=1, row=1, sticky=tk.EW, padx=5, pady=5)

# broj prijave
broj_label = ttk.Label(lf2, text="Broj prijave")
broj_label.grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)

broj_entry = ttk.Entry(lf2)
broj_entry.grid(column=1, row=2, sticky=tk.EW, padx=5, pady=5)



def stavi_stambilj():
    ########  UNOS DATUMA I BROJA PREDMETA NA STAMBILJ #######
    # Open an Image
    try:
        img = Image.open('slika.png')

        # Custom font style and font size
        myFont1 = ImageFont.truetype('arial.ttf', 62)
        myFont2 = ImageFont.truetype('arial.ttf', 72)

        # Call draw Method to add 2D graphics in an image
        I1 = ImageDraw.Draw(img)

        datum = str(datum_entry.get())
        broj = str(broj_entry.get())

        # Add Text to an image
        I1.text((350, 173), datum, font=myFont1, fill=(7, 60, 145))
        I1.text((30, 340), broj, font=myFont2, fill=(7, 60, 145))

        # Display edited image
        # img.show()

        izmenjen_broj = broj.replace("/", "-")

        # Save the edited image
        img.save("stambilj.png")

        ############## STAVLJANJE STAMBILJA NA PDF PRIJAVU ####################

        input_file = str(proba_pdf)
        output_file = "prijava" + izmenjen_broj + ".pdf"
        barcode_file = "stambilj.png"

        # define the position (upper-right corner)
        image_rectangle = fitz.Rect(350, 20, 600, 100)

        # retrieve the first page of the PDF
        file_handle = fitz.open(input_file)
        first_page = file_handle[0]

        # add the image
        first_page.insert_image(image_rectangle, filename=barcode_file)

        # brisanje iskorišćenog stambilja
        if os.path.exists("stambilj.png"):
            os.remove("stambilj.png")
        else:
            print("The file does not exist")

        radna_povrsina = os.path.join((os.environ['USERPROFILE']), 'Desktop')

        overena_prijava = os.path.join(radna_povrsina, output_file)

        file_handle.save(overena_prijava)

        info_poruka.set("Štambilj je stavljen na dokument!")
    except:
        info_poruka.set("Nešto nije u redu, pokušajte ponovo.")



# Završi button
zavrsi_button = ttk.Button(lf2, text="Završi", command=stavi_stambilj)
zavrsi_button.grid(column=0, row=3, columnspan = 2, sticky=tk.EW, padx=5, pady=5)



root.mainloop()





