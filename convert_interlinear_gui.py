import tkinter as tk
from tkinter import ttk, filedialog
import os
from xml.etree.ElementTree import tostring
from xml.dom import minidom

from excel_to_xml import convert_excel_to_xml_dom
from xml_to_flextext import transform_to_flextext_dom


# TODO refactor as a class

def main():
    global intermediate_xml
    intermediate_xml = None

    def choose_input_file():
        """
        Open a file dialog to choose an input file based on the selected format.
        """

        formatString = inputFormatCombo.get()
        filetypelist = [("Excel files", "*.xlsx *.xls"),
                        ("All files", "*.*")]
        filepath = filedialog.askopenfilename(title="Select a file", filetypes=filetypelist)
        if filepath:
            # check whether the file exists and is readable...
            pass
        _, inputExt = os.path.splitext(filepath)
        if inputExt.lower() in [".xlsx", ".xls"]:
            inputFormatCombo.set("Excel Interlinear")
        inputFileName.config(text=filepath)
        inputLoadButton.state(['!disabled'])
        outputFileName.config(text="")
        convertButton.state(['disabled'])

    def choose_output_file():
        """
        Open a file dialog to choose an output file based on the selected format.
        """

        filetypelist = [("FlexText files", "*.flextext"),
                        ("All files", "*.*")]
        filenamebase, _ = os.path.splitext(inputFileName.cget("text"))
        initialfilename = filenamebase + ".flextext"
        filepath = filedialog.asksaveasfilename(initialfile=initialfilename, defaultextension=".flextext")
        if filepath:
            # check whether the file can be created/written...
            pass
        _, outputExt = os.path.splitext(filepath)
        if outputExt.lower() in [".flextext"]:
            outputFormatCombo.set("FlexText Interlinear")
        outputFileName.config(text=filepath)
        convertButton.state(['!disabled'])

    def load_file():
        """
        Load the selected input file and convert it to an intermediate XML format.
        """

        global intermediate_xml

        # Should this happen automatically after choosing a file? Maybe later.
        filepath = inputFileName.cget("text")
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

        # global intermediate_xml

        (flextext_xml, missing_freetrans_count) = transform_to_flextext_dom(
            intermediate_xml, wsVernacular.get(), wsGloss.get(), wsFree.get())
        pretty_xml = prettify_xml(flextext_xml)

        with open(outputFileName.get(), 'w', encoding='utf-8') as f:
            f.write(pretty_xml)

    def prettify_xml(element):
        """
        Return a pretty-printed XML string for the given element.
        """

        rough_string = tostring(element, encoding='utf-8')
        reparsed = minidom.parseString(rough_string)
        # Ensure UTF-8 output
        return reparsed.toprettyxml(encoding='utf-8', indent="  ").decode('utf-8')

    # main code

    root = tk.Tk()
    root.title("Interlinear Converter")
    # root.geometry("400x300")

    mainframe = ttk.Frame(root, padding="10 10 10 10")
    mainframe.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))
    
    # Input block
    inputFileName = ttk.Label(mainframe, text="Input file name", width=30)
    inputFileName.grid(row=0, column=0, columnspan=2, pady=5, padx=5)
    inputBrowseButton = ttk.Button(mainframe, text="Choose File ...", command=choose_input_file)
    inputBrowseButton.grid(row=0, column=2, pady=5, padx=5)
    inputFormatLabel = ttk.Label(mainframe, text="Input Format:")
    inputFormatLabel.grid(row=1, column=0, pady=5, padx=5)
    inputFormatCombo = ttk.Combobox(mainframe, values=["Excel Interlinear"])
    inputFormatCombo.grid(row=1, column=1, pady=5, padx=5)
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
    outputFileName = ttk.Label(mainframe, text="Output file name", width=30)
    outputFileName.grid(row=7, column=0, columnspan=2, pady=5, padx=5)
    outputBrowseButton = ttk.Button(mainframe, text="Choose File ...", command=choose_output_file)
    outputBrowseButton.grid(row=7, column=2, pady=5, padx=5)
    outputFormatLabel = ttk.Label(mainframe, text="Output Format:")
    outputFormatLabel.grid(row=8, column=0, pady=5, padx=5)
    outputFormatCombo = ttk.Combobox(mainframe, values=["FlexText Interlinear"])
    outputFormatCombo.grid(row=8, column=1, pady=5, padx=5)
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