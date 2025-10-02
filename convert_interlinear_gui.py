import tkinter as tk
from tkinter import ttk
from excel_to_xml import convert_excel_to_xml_dom
from xml_to_flextext import transform_to_flextext_dom

global intermediate_xml # TEMP: to pass XML data from load_file() to convert()
# TODO refactor as a class

def main():
    def choose_input_file():
        """
        Open a file dialog to choose an input file based on the selected format.
        """

        formatString = inputFormatCombo.get()
        if formatString == "Excel Interlinear":
            filetypelist = [("Excel files", "*.xlsx")]
        # elif ...
        else:
            # show error: need a format selected
            pass
        filetypelist.append(("All files", "*.*"))
        filepath = tk.filedialog.askopenfilename(title="Select a file", filetypes=filetypelist)
        if filepath:
            # check whether the file exists and is readable...
            pass
        inputFileName.text = filepath
        inputLoadButton.state(['!disabled'])

    def load_file():
        """
        Load the selected input file and convert it to an intermediate XML format.
        """

        # Should this happen automatically after choosing a file? Maybe later.
        filepath = inputFileName.text
        formatString = inputFormatCombo.get()
        if formatString == "Excel Interlinear":
            # utilize excel-to-xml code
            (intermediate_xml, error_list) = convert_excel_to_xml_dom(filepath)
            # set inputLang and inputWritingSystem labels
            # inputLang.config(text="Language: [set after loading file]")

    def convert():
        """
        Convert intermediate XML to FlexText interlinear format
        """

        (flextext_xml, missing_freetrans_count) = transform_to_flextext_dom(intermediate_xml,
                                                                            wsVernacular.get(), wsGloss.get(), wsFree.get())

    # main code

    root = tk.Tk()
    root.title("Interlinear Converter")
    # root.geometry("400x300")

    mainframe = ttk.Frame(root, padding="10 10 10 10")
    mainframe.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))
    
    # Input block
    inputFormatLabel = ttk.Label(mainframe, text="Input Format:")
    inputFormatLabel.grid(row=0, column=0, pady=5, padx=5)
    inputFormatCombo = ttk.Combobox(mainframe, values=["Excel Interlinear"])
    inputFormatCombo.grid(row=0, column=1, pady=5, padx=5)

    inputFileName = ttk.Label(mainframe, text="Input file name", width=30)
    inputFileName.grid(row=1, column=0, pady=5, padx=5)
    inputBrowseButton = ttk.Button(mainframe, text="Choose File ...", command=choose_input_file)
    inputBrowseButton.grid(row=1, column=1, pady=5, padx=5)
    inputLoadButton = ttk.Button(mainframe, text="Load File", state='disabled', command=load_file)
    inputLoadButton.grid(row=1, column=2, pady=5, padx=5)
    
    # Reminders & writing systems
    reminderLabel = ttk.Label(mainframe, text="Reminders about setting language and writing systems in Excel and in FLEx.")
    reminderLabel.grid(row=3, column=0, columnspan=3, pady=5, padx=5)
    reminderLabel.config(wraplength=400)
    wsVernacularLabel = ttk.Label(mainframe, text="Vernacular writing system:", anchor='e')
    wsVernacularLabel.grid(row=4, column=0, pady=5, padx=5)
    wsVernacular = ttk.Entry(mainframe, text="")   # set default after loading file
    wsVernacular.grid(row=4, column=1, pady=5, padx=5)
    wsFreeLabel = ttk.Label(mainframe, text="Free trans. writing system:", anchor='e')
    wsFreeLabel.grid(row=5, column=0, pady=5, padx=5)
    wsFree = ttk.Entry(mainframe, text="")   # set default after loading file
    wsFree.grid(row=5, column=1, pady=5, padx=5)
    wsGlossLabel = ttk.Label(mainframe, text="Gloss writing system:", anchor='e')
    wsGlossLabel.grid(row=6, column=0, pady=5, padx=5)
    wsGloss = ttk.Entry(mainframe, text="")   # set default after loading file
    wsGloss.grid(row=6, column=1, pady=5, padx=5)

    # Output block
    outputFormatLabel = ttk.Label(mainframe, text="Output Format:")
    outputFormatLabel.grid(row=7, column=0, pady=5, padx=5)
    outputFormatCombo = ttk.Combobox(mainframe, values=["FlexText Interlinear"])
    outputFormatCombo.grid(row=7, column=1, pady=5, padx=5)
    outputFileName = ttk.Entry(mainframe, width=30)
    outputFileName.grid(row=8, column=0, pady=5, padx=5)
    outputBrowseButton = ttk.Button(mainframe, text="Choose File ...")
    outputBrowseButton.grid(row=8, column=1, pady=5, padx=5)
    convertButton = ttk.Button(mainframe, text="Convert", state='disabled', command=convert)
    convertButton.grid(row=8, column=2, pady=5, padx=5)

    # Error display
    errorDisplay = tk.Text(mainframe, height=5, width=50)
    errorDisplay.grid(row=9, column=0, columnspan=3, pady=5, padx=5)
    errorDisplay.config(state=tk.DISABLED)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()

if __name__ == "__main__":
    main()