from tkinter import *
import pandas as pd

root = Tk()
root.title("Registers")
frame = Frame(root)
frame.pack()
label = Label(root, fg="dark green")
label.pack()


def ERF():
    ERF_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Emissions%20Reduction%20Fund%20Register.csv"
    ERF = pd.read_csv(ERF_url, encoding="ISO-8859-1")
    return ERF.to_excel('ERF Register.xlsx')


def CAC():
    CAC_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Carbon%20Abatement%20Contract%20table.csv"
    CAC = pd.read_csv(CAC_url, encoding="ISO-8859-1")
    return CAC.to_excel('CAC Register.xlsx')


def Vol():
    Vol_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Voluntary%20Cancellations.csv"
    Vol = pd.read_csv(Vol_url, encoding="ISO-8859-1")
    return Vol.to_excel('Voluntary Register.xlsx')


ERF_button = Button(frame, text="ERF Register", fg="dark green", command=ERF)
ERF_button.pack(fill=BOTH, expand=True)

CAC_button = Button(frame, text="CAC Register", fg="dark green", command=CAC)
CAC_button.pack(fill=BOTH, expand=True)

Vol_button = Button(frame, text="Voluntary Register", fg="dark green", command=Vol)
Vol_button.pack(fill=BOTH, expand=True)

# Poi aggiungo qui l'opzione del monthly report.

quit_button = Button(frame, text="QUIT", fg="red", command=quit)
quit_button.pack(fill=BOTH, expand=True)





root.mainloop()