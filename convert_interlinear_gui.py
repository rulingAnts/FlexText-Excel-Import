import tkinter as tk
from tkinter import ttk, filedialog
import os
from xml.etree.ElementTree import tostring
from xml.dom import minidom

from excel_to_xml import convert_excel_to_xml_dom
from xml_to_flextext import transform_to_flextext_dom



class Converter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Interlinear Converter")
        self.intermediate_xml = None

        self.mainframe = ttk.Frame(self, padding="10 10 10 10")
        self.mainframe.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))

        # Input block
        self.inputFileName = ttk.Label(self.mainframe, text="Input file name", width=30)
        self.inputFileName.grid(row=0, column=0, columnspan=2, pady=5, padx=5)
        self.inputBrowseButton = ttk.Button(self.mainframe, text="Choose File ...", command=self.choose_input_file)
        self.inputBrowseButton.grid(row=0, column=2, pady=5, padx=5)
        self.inputFormatLabel = ttk.Label(self.mainframe, text="Input Format:")
        self.inputFormatLabel.grid(row=1, column=0, pady=5, padx=5)
        self.inputFormatCombo = ttk.Combobox(self.mainframe, values=["Excel Interlinear"])
        self.inputFormatCombo.grid(row=1, column=1, pady=5, padx=5)
        self.inputLoadButton = ttk.Button(self.mainframe, text="Load File", state='disabled', command=self.load_file)
        self.inputLoadButton.grid(row=1, column=2, pady=5, padx=5)

        # Reminders & writing systems
        self.reminderLabel = ttk.Label(self.mainframe, text="Reminders about setting language and writing systems in Excel and in FLEx.")
        self.reminderLabel.grid(row=3, column=0, columnspan=3, pady=5, padx=5)
        self.reminderLabel.config(wraplength=400)
        self.wsVernacularLabel = ttk.Label(self.mainframe, text="Vernacular writing system:", anchor='e')
        self.wsVernacularLabel.grid(row=4, column=0, pady=5, padx=5)
        self.wsVernacular = ttk.Entry(self.mainframe)
        self.wsVernacular.grid(row=4, column=1, pady=5, padx=5)
        self.wsFreeLabel = ttk.Label(self.mainframe, text="Free trans. writing system:", anchor='e')
        self.wsFreeLabel.grid(row=5, column=0, pady=5, padx=5)
        self.wsFree = ttk.Entry(self.mainframe)
        self.wsFree.grid(row=5, column=1, pady=5, padx=5)
        self.wsGlossLabel = ttk.Label(self.mainframe, text="Gloss writing system:", anchor='e')
        self.wsGlossLabel.grid(row=6, column=0, pady=5, padx=5)
        self.wsGloss = ttk.Entry(self.mainframe)
        self.wsGloss.grid(row=6, column=1, pady=5, padx=5)

        # Output block
        self.outputFileName = ttk.Label(self.mainframe, text="Output file name", width=30)
        self.outputFileName.grid(row=7, column=0, columnspan=2, pady=5, padx=5)
        self.outputBrowseButton = ttk.Button(self.mainframe, text="Choose File ...", command=self.choose_output_file)
        self.outputBrowseButton.grid(row=7, column=2, pady=5, padx=5)
        self.outputFormatLabel = ttk.Label(self.mainframe, text="Output Format:")
        self.outputFormatLabel.grid(row=8, column=0, pady=5, padx=5)
        self.outputFormatCombo = ttk.Combobox(self.mainframe, values=["FlexText Interlinear"])
        self.outputFormatCombo.grid(row=8, column=1, pady=5, padx=5)
        self.convertButton = ttk.Button(self.mainframe, text="Convert", state='disabled', command=self.convert)
        self.convertButton.grid(row=8, column=2, pady=5, padx=5)

        # Error display
        self.errorDisplay = tk.Text(self.mainframe, height=5, width=50)
        self.errorDisplay.grid(row=9, column=0, columnspan=3, pady=5, padx=5)
        self.errorDisplay.config(state=tk.DISABLED)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    def choose_input_file(self):
        """
        Opens a file dialog to choose an input file.
        """

        formatString = self.inputFormatCombo.get()
        filetypelist = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        filepath = filedialog.askopenfilename(title="Select a file", filetypes=filetypelist)
        if filepath:
            _, inputExt = os.path.splitext(filepath)
            if inputExt.lower() in [".xlsx", ".xls"]:
                self.inputFormatCombo.set("Excel Interlinear")
            self.inputFileName.config(text=filepath)
            self.inputLoadButton.state(['!disabled'])
            self.outputFileName.config(text="")
            self.convertButton.state(['disabled'])

    def choose_output_file(self):
        """
        Opens a file dialog to choose an output file.
        """

        filetypelist = [("FlexText files", "*.flextext"), ("All files", "*.*")]
        filenamebase, _ = os.path.splitext(self.inputFileName.cget("text"))
        initialfilename = filenamebase + ".flextext"
        filepath = filedialog.asksaveasfilename(initialfile=initialfilename, defaultextension=".flextext")
        if filepath:
            _, outputExt = os.path.splitext(filepath)
            if outputExt.lower() in [".flextext"]:
                self.outputFormatCombo.set("FlexText Interlinear")
            self.outputFileName.config(text=filepath)
            self.convertButton.state(['!disabled'])

    def load_file(self):
        """
        Loads the input file and processes it into intermediate XML.
        """

        filepath = self.inputFileName.cget("text")
        formatString = self.inputFormatCombo.get()
        if formatString == "Excel Interlinear":
            (self.intermediate_xml, error_list) = convert_excel_to_xml_dom(filepath)
            # TODO: set inputLang and inputWritingSystem labels
            # self.inputLang.config(text="Language: [set after loading file]")

    def convert(self):
        """
        Converts the intermediate XML into FlexText format.
        """

        (flextext_xml, missing_freetrans_count) = transform_to_flextext_dom(
            self.intermediate_xml, self.wsVernacular.get(), self.wsGloss.get(), self.wsFree.get())
        pretty_xml = self.prettify_xml(flextext_xml)
        with open(self.outputFileName.cget("text"), 'w', encoding='utf-8') as f:
            f.write(pretty_xml)

    def prettify_xml(self, element):
        """
        Add whitespace to make an XML element tree pretty-printed.
        """

        rough_string = tostring(element, encoding='utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(encoding='utf-8', indent="  ").decode('utf-8')


if __name__ == "__main__":
    app = Converter()
    app.mainloop()