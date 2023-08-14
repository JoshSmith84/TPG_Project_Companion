# companion_app.py
# Author:Josh Smith

import sys
from pathlib import Path
import csv
import tkinter as tk
from tkinter import ttk


class BoundText(tk.Text):
    """A Text widget with a bound variable
    Credit: Alan D. Moore "Python GUI Programming with Tkinter"""

    def __init__(self, *args, textvariable=None, **kwargs):
        super().__init__(*args, **kwargs)
        self._variable = textvariable
        if self._variable:
            self.insert('1.0', self._variable.get())
            self.bind('<<Modified>>', self._set_var)

    def _set_content(self, *_):
        """Set the text contents to the variable"""
        self.delete('1.0', tk.END)
        self.insert('1.0', self._variable.get())

    def _set_var(self, *_):
        """Set the variable to the text contents"""
        if self.edit_modified():
            content = self.get('1.0', 'end-1chars')
            self._variable.set(content)
            self.edit_modified(False)


class LabelInput(tk.Frame):
    """A widget containing a label and input together.
    Credit: Alan D. Moore "Python GUI Programming with Tkinter"""

    def __init__(
            self, parent, label, var, input_class=ttk.Entry,
            input_args=None, label_args=None, **kwargs
    ):
        super().__init__(parent, **kwargs)
        input_args = input_args or {}
        label_args = label_args or {}
        # The above statements say if label_args or input_args are not None,
        # they are what was passed during init.
        # However if they are None, then make them empty dicts
        self.variable = var
        self.variable.label_widget = self

        if input_class in (ttk.Checkbutton, ttk.Button):
            input_args["text"] = label
        else:
            self.label = ttk.Label(self, text=label, **label_args)
            self.label.grid(row=0, column=0, sticky=(tk.W + tk.E))

        if input_class in (
            ttk.Checkbutton, ttk.Button, ttk.Radiobutton
        ):
            input_args["variable"] = self.variable
        else:
            input_args["textvariable"] = self.variable

        if input_class == ttk.Radiobutton:
            self.input = tk.Frame(self)
            for v in input_args.pop('values', []):
                button = ttk.Radiobutton(
                    self.input, value=v, text=v, **input_args
                )
                button.pack(
                    side=tk.LEFT, ipadx=10, ipady=2, expand=True, fill='x'
                )
        else:
            self.input = input_class(self, **input_args)

        self.input.grid(row=1, column=0, sticky=(tk.E + tk.W))
        self.columnconfigure(0, weight=1)

    def grid(self, sticky=(tk.E + tk.W), **kwargs):
        """Override grid to add default sticky values"""
        super().grid(sticky=sticky, **kwargs)


class AppPage(ttk.Frame):
    """Application page class from which all other pages will inherit."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._vars = {
            'Client Name': tk.StringVar(),
            'Go Live / Dead Date': tk.StringVar(),
            'Market': tk.StringVar(),
            'Type': tk.StringVar(),
            'AV': tk.StringVar(),
            'MDR': tk.StringVar(),
            'Ninjio': tk.StringVar(),
            'Barracuda': tk.StringVar(),
        }

    def _add_frame(self, label, cols=2):
        """Add a label frame to the form
        Credit: Alan D. Moore "Python GUI Programming with Tkinter"""

        frame = ttk.LabelFrame(self, text=label)
        frame.grid(sticky=tk.W + tk.E)
        for i in range(cols):
            frame.columnconfigure(i, weight=1)
        return frame

    def get(self):
        data = dict()
        for key, variable in self._vars.items():
            try:
                data[key] = variable.get()
            except tk.TclError:
                message = f'Error in field: {key}. Data was not saved!'
                raise ValueError(message)
        return data

    def _tpg_tools(self, p_type, column):
        """Lays out the frames needed to select TPG tools.
        This is needed for onbaordings(to see what to deploy)
        Offboardings(to see what to pull)
        Conversiones(to see what to pull before push or just review)"""

        if p_type == 'Offboarding':
            av_info = self._add_frame(f"Which AV are we pulling?")
            mdr_info = self._add_frame(f"Which MDR are we pulling?")
            ess_info = self._add_frame(
                f"Which Barracuda level are we pulling?")
        elif p_type == 'Conversion':
            av_info = self._add_frame(f"Which AV do they currently have?")
            mdr_info = self._add_frame(f"Which MDR do they currently have?")
            ess_info = self._add_frame(
                f"Which Barracuda level do they currently have?")
        else:
            av_info = self._add_frame(f"Which AV are they getting?")
            mdr_info = self._add_frame(f"Which MDR are they getting?")
            ess_info = self._add_frame(
                f"Which Barracuda level are they getting?")

        LabelInput(
            av_info, "", input_class=ttk.Radiobutton,
            var=self._vars['AV'],
            input_args={"values": ['Sophos', 'AV Defender', 'None']}
        ).grid(row=0, column=column)

        LabelInput(
            mdr_info, "", input_class=ttk.Radiobutton,
            var=self._vars['MDR'],
            input_args={"values": ['Blackpoint', 'Arctic Wolf', 'None']}
        ).grid(row=0, column=column)

        LabelInput(
            ess_info, "", input_class=ttk.Radiobutton,
            var=self._vars['Barracuda'],
            input_args={"values": ['Complete+Sentinel',
                                   'Complete Only',
                                   'Spam only',
                                   'None'
                                   ]
                        }
        ).grid(row=0, column=column)

        ninjio_info = self._add_frame("Security Training?")
        LabelInput(
            ninjio_info, "Ninjio?", input_class=ttk.Checkbutton,
            var=self._vars['Ninjio'],
        ).grid(row=0, column=column)


    @staticmethod
    def _on_quit():
        """Command to exit program"""
        sys.exit()


class MainPage(AppPage):
    """Main page to choose whether to create, update, or run report."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        m_menu = self._add_frame("Main Menu")
        self.create_button = ttk.Button(
            m_menu,
            text='Create Project',
            command=self._on_create
        )
        self.create_button.grid(row=0, column=0, sticky='ew')

        self.update_button = ttk.Button(
            m_menu,
            text='Update Project',
            command=self._on_update
        )
        self.update_button.grid(row=1, column=0, sticky='ew')

        self.report_button = ttk.Button(
            m_menu,
            text='Run Report',
            command=self._on_report
        )
        self.report_button.grid(row=2, column=0, sticky='ew')

        self.quit_button = ttk.Button(
            m_menu,
            text='Quit',
            command=self._on_quit
        )
        self.quit_button.grid(row=3, column=0, sticky='ew')

    @staticmethod
    def _on_create():
        """Command to launch new project page"""
        app.create_page1()

    def _on_update(self):
        """Command to update existing project"""
        pass

    def _on_report(self):
        """Command to open report page"""
        pass


# noinspection PyTypeChecker
class CreatePage1(AppPage):
    """Project create page #1"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        c_info = self._add_frame("Client Information")

        LabelInput(
            c_info, "Client Name", var=self._vars['Client Name']
        ).grid(row=0, column=0)
        LabelInput(
            c_info, "Go Live / Dead Date",
            var=self._vars['Go Live / Dead Date']
        ).grid(row=1, column=0)

        m_info = self._add_frame('Market')

        LabelInput(
            m_info, "", input_class=ttk.Radiobutton,
            var=self._vars['Market'],
            input_args={"values": ["NO", "SHV", "ATX", "IND", "VAL", "KC"]},
        ).grid(row=0, column=0)

        t_info = self._add_frame('Project Type')

        LabelInput(
            t_info, "", input_class=ttk.Radiobutton,
            var=self._vars['Type'],
            input_args={"values": ["Onboarding", "Offboarding", "Conversion"]}
        ).grid(row=0, column=0)

        buttons = tk.Frame(self)
        buttons.grid(sticky=tk.W + tk.E + tk.S, row=3)

        self.quit_button = ttk.Button(
            buttons, text="Quit", command=self._on_quit
        )
        self.quit_button.pack(side=tk.RIGHT)
        self.next_button = ttk.Button(
            buttons, text="Next", command=self._on_create2
        )
        self.next_button.pack(side=tk.RIGHT)

    def _on_create2(self):
        """Move onto creation page two
        (which will be one of three pages depending on Type selected"""
        app.create_page2()


# noinspection PyTypeChecker
class CreatePage2On(AppPage):
    """Creation Page 2 if 'Onboardings' is type"""

    def __init__(self, client, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.client_label = ttk.Label(
            self,
            text=f"{client}'s Onboarding",
            font=("TKDefaultFont", 12))
        self.client_label.grid(row=0)

        self._tpg_tools('Onboarding', 0)

        buttons = tk.Frame(self)
        buttons.grid(sticky=tk.W + tk.E + tk.S, row=5)

        self.quit_button = ttk.Button(
            buttons, text="Quit", command=self._on_quit
        )
        self.quit_button.pack(side=tk.RIGHT)
        self.next_button = ttk.Button(
            buttons, text="Save", command=self._on_create2on_save
        )
        self.next_button.pack(side=tk.RIGHT)

    def _on_create2on_save(self):
        app.creation_on_complete()


class CreatePage2Off(AppPage):
    """Creation Page 2 if 'Offboardings' is type"""

    def __init__(self, client, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.client_label = ttk.Label(
            self,
            text=f"{client}'s Offboarding",
            font=("TKDefaultFont", 12))
        self.client_label.grid(row=0)

        self._tpg_tools('Offboarding', 0)

        buttons = tk.Frame(self)
        buttons.grid(sticky=tk.W + tk.E + tk.S, row=5)

        self.quit_button = ttk.Button(
            buttons, text="Quit", command=self._on_quit
        )
        self.quit_button.pack(side=tk.RIGHT)
        self.next_button = ttk.Button(
            buttons, text="Save", command=self._on_createoff_save
        )
        self.next_button.pack(side=tk.RIGHT)

    def _on_createoff_save(self):
        pass

class CreatePage2Conv(AppPage):
    """Creation Page 2 if 'Conversions' is type"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)


class Application(tk.Tk):
    """Application root window"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("TPG Onboarding Project Companion")
        self.columnconfigure(0, weight=1)
        self.minsize(400, 300)
        self.main_label = ''
        self.m_page = ''
        self.c_page1 = CreatePage1(self)
        self.c_page2_on = CreatePage2On(self)
        self.c_page2_off = CreatePage2Off(self)
        self.client_data = {}
        self.new_plan_data = {}

        self.status = tk.StringVar()
        ttk.Label(self, textvariable=self.status
                  ).grid(sticky=(tk.W + tk.E), row=2, padx=10)
        self.main_page()

    def main_page(self):
        self.m_page = MainPage(self)
        self.main_label = ttk.Label(
            self,
            text="TPG Onboarding Project Companion",
            font=("TKDefaultFont", 14))
        self.main_label.grid(row=0)
        self.m_page.grid(row=1, padx=10, sticky=(tk.W + tk.E))

    def create_page1(self):
        self.main_label.grid_forget()
        self.main_label.destroy()
        self.m_page.grid_forget()
        self.m_page.destroy()
        self.c_page1 = CreatePage1(self)
        self.c_page1.grid(row=1, padx=10, sticky=(tk.W + tk.E))

    def create_page2(self):
        try:
            self.client_data = self.c_page1.get()
        except ValueError as e:
            self.status.set(str(e))
            return
        client_name = self.client_data['Client Name']

        if self.client_data['Market'] == '':
            self.status.set('Required Market Location is Missing!')

        self.c_page1.grid_forget()
        self.c_page1.destroy()

        if self.client_data['Type'] == 'Onboarding':
            self.c_page2_on = CreatePage2On(client_name)
            self.c_page2_on.grid(row=1, padx=10, sticky=(tk.W + tk.E))
        elif self.client_data['Type'] == 'Offboarding':
            self.c_page2_off = CreatePage2Off(client_name)
            self.c_page2_off.grid(row=1, padx=10, sticky=(tk.W + tk. E))
        elif self.client_data['Type'] == 'Conversion':
            print('open Conversion page')
        else:
            self.status.set('Required Project Type is Missing!')

    def creation_on_complete(self):
        try:
            self.new_plan_data = self.c_page2_on.get()
        except ValueError as e:
            self.status.set(str(e))
            return
        print(self.client_data)
        print(self.new_plan_data)
        self.c_page2_on.grid_forget()
        self.c_page2_on.destroy()
        self.main_page()


if __name__ == "__main__":
    app = Application()
    app.mainloop()