from tkinter import *
import disposal, vials, tc99mgen, ga68gen, i131
from functions import *

class Tabs:

    def __init__(self, window, tabs_frame, main_tab):
        self.window = window
        self.tabs_frame = tabs_frame
        self.main_tab = main_tab

    def back_to_main(self, tab_name):
        self.tabs_frame.forget(tab_name)
        self.tabs_frame.select(self.main_tab)

    #=====CREATE TABS=====
    def create_new_tab(self, tab_name):
        for tab in self.tabs_frame.tabs():
            if self.tabs_frame.tab(tab, "text") == tab_name:
                self.tabs_frame.select(tab)
                return
        new_tab = Frame(self.tabs_frame, bg=C4)
        self.tabs_frame.add(new_tab, text=tab_name)
        self.tabs_frame.select(new_tab)
        if tab_name == "Vials":
            self._tab_vials(new_tab)
        elif tab_name == "Generators":
            self._tab_generators(new_tab)
        elif tab_name == "I131":
            i131.build_tab(self, new_tab)
        elif tab_name == "Tc99m Gen":
            tc99mgen.build_tab(self, new_tab)
        elif tab_name == "Ga68 Gen":
            ga68gen.build_tab(self, new_tab)
        elif tab_name in [name for name, _ in VIAL_DATA]:
            vials.build_tab(self, new_tab, tab_name)
        elif tab_name == "Disposal":
            self._tab_disposal(new_tab)

    #=====CUSTOMIZE EACH TAB=====
    def _tab_vials(self, tab):
        content_frame = Frame(tab, bg=C4)
        content_frame.place(relx=0.5, rely=0.5, anchor="center")
        Label(content_frame, text="Select New Radionuclide:", bg=C4, fg="white", font=(FONT_NAME, 24, "bold")).grid(columnspan=3, column=0, row=0, pady=(0, 20))
        for idx, (name,_) in enumerate(VIAL_DATA):
            r = idx // 3 + 1
            c = idx % 3
            Button(content_frame, text=name, **TAB_BUTTON_STYLE, command=lambda t=name: self.create_new_tab(t)).grid(column=c, row=r, padx=10, pady=10)
        Button(content_frame, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: self.back_to_main(nt)).grid(columnspan=3, column=0, row=5, pady=(30, 0))

    def _tab_generators(self, tab):
        frame = Frame(tab, bg=C4)
        frame.pack(expand=True)
        Button(frame, text="Tc99m Gen", **GEN_BUTTON_STYLE, command=lambda: self.create_new_tab("Tc99m Gen")).grid(row=0, column=0,sticky="e", padx=20, pady=10)
        Button(frame, text="Ga68 Gen", **GEN_BUTTON_STYLE, command=lambda : self.create_new_tab("Ga68 Gen")).grid(row=0, column=1, sticky="w", padx=20, pady=10)
        Button(tab, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: self.back_to_main(nt)).pack(pady=40)

    def _tab_disposal(self, tab):
        disposal.build_disposal_tab(tab, on_back=lambda: self.back_to_main(tab))