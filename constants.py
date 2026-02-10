import math

#=====Customization=====}
C1, C2, C3, C4, BG, FONT_NAME = "#b2d8d8", "#66b2b2", "#008080", "#006666", "#004c4c", "Times"
TEXT_COLORS = {"bg" : C4, "fg" : "white"}
TEXT_COLORS_KITS ={"bg": C3, "fg": "white"}
BUTTON_STYLE = {
    "background": C3,
    "foreground": "white",
    "activebackground": C3,
    "activeforeground": "white",
    "width": 14,
    "height": 2,
    "cursor": "hand2",
    "highlightthickness": 0,
    "font": (FONT_NAME, 16, "bold")
}
TAB_BUTTON_STYLE = {
    "background": BG,
    "foreground": "white",
    "activebackground": BG,
    "activeforeground": "white",
    "width": 12,
    "height": 2,
    "cursor": "hand2",
    "highlightthickness": 0,
    "font": (FONT_NAME, 14, "bold")
}
GEN_BUTTON_STYLE = {
    "bg" : BG,
    "fg" : "white",
    "width" : 10,
    "height" : 2,
    "font" : (FONT_NAME,30,"bold"),
    "activebackground" : "white",
    "activeforeground" : BG,
    "cursor" : "hand2",
    "highlightthickness" : 0,
}

# =====Τ1/2=====
T12_I131, T12_GE68 = 8.0207, 271.05 #days
T12_TC99M, T12_MO99 = 6.015, 65.94 #hours
LAMBDA_MO = math.log(2) / T12_MO99
LAMBDA_TC = math.log(2) / T12_TC99M
T12_GA68 = 67.84 #minutes

#=====Vials=====
VIAL_DATA = [("51-Cr",664.86), ("59-Fe",1067.9), ("67-Ga",78.3), ("89-Sr",1212.7), ("90-Y",64.1), ("111-In",67.3),
             ("123-I",13.27), ("125-I",1425.6), ("131-I",192.5), ("153-Sm",46.5), ("186-Re",89.2), ("201-Tl",72.9)]

#=====Disposal=====
DISPOSAL_LIMITS_BQ = {"51-Cr": 1e7, "59-Fe": 1e6, "89-Sr": 1e6, "90-Y":  1e5, "111-In": 1e6, "123-I": 1e7,
                      "125-I": 1e6, "131-I": 1e6, "153-Sm": 1e6, "186-Re": 1e6, "201-Tl": 1e6,}

#=====Kits=====
KIT_CONFIG = {
    "DTPA": {
        "title": "Lung Ventilation",
        "default_activity": 50,
        "dilution": "5ml",
        "shake": "2min",
        "final_volume": 5
    },
    "MDP": {
        "title": "Bone Scan",
        "default_activity": 150,
        "dilution": "5ml",
        "shake": "30sec",
        "store": "to fridge",
        "final_volume": 5
    },
    "MAG-3": {
        "title": "Dynamic Kidney Study",
        "default_activity": 30,
        "dilution": "10ml",
        "boil": "10min",
        "cooling": "under the tab for 10min",
        "store": "to fridge",
        "final_volume": 10
    },
    "MAASCINT": {
        "title": "Pulmonary Perfusion",
        "default_activity": 40,
        "dilution": "4ml",
        "shake": "few times",
        "incubation": "5min",
        "final_volume": 4
    },
    "DMSA": {
        "title": "Static Kidney Study",
        "default_activity": 20,
        "dilution": "5ml",
        "remove": "equal volume of air",
        "shake": "5-10min",
        "incubation": "5-10min",
        "final_volume": 5
    },
    "MYOVIEW": {
        "title": "Myocardial Perfusion",
        "default_activity": 250,
        "dilution": "8ml",
        "needs": "breathing needle",
        "shake": "few times",
        "incubation": "15min",
        "use": "Tc99m eluted within 6h",
        "final_volume": 8
    },
    "PHYTATE": {
        "title": "Liver",
        "default_activity": 40,
        "dilution": "5ml",
        "shake": "2min",
        "store": "to fridge",
        "final_volume": 5
    },
    "HIG": {
        "title": "Immunoscintigraphy",
        "default_activity": 15,
        "dilution": "= req vol",
        "remove": "equal volume of air",
        "incubation": "20min",
        "dilution with saline": "after the incubation",
        "final_volume": "= req vol"
    },
    "CERETEC": {
        "title": "Brain",
        "default_activity": 40,
        "dilution": "5ml",
        "shake": "10sec",
        "use": "Tc99m eluted within 2h",
        "final_volume": 5
    },
    "CEA-SCAN": {
        "title": "CEA-SCAN",
        "default_activity": 25,
        "dilution": "1ml",
        "shake": "30sec",
        "incubation": "5min",
        "final_volume": 1
    },
    "LEUKOSCAN": {
        "title": "LEUKOSCAN",
        "default_activity": 27,
        "dilution": "1ml",
        "add 0.5ml saline": "shake 30sec",
        "add": "Tc99m",
        "shake": "few times",
        "incubation": "5min",
        "final_volume": 1.5
    },
    "BIDA": {
        "title": "Bile",
        "default_activity": 20,
        "dilution": "5ml",
        "shake": "few times",
        "incubation": "15min",
        "shake before": "use",
        "final_volume": 5
    },
    "CARDIOLITE": {
        "title": "Myocardial Perfusion",
        "default_activity": 50,
        "dilution": "5ml",
        "shake": "few times",
        "boil": "10-12min",
        "cooling": "15min",
        "store": "to fridge",
        "final_volume": 5
    },
    "NEUROLITE": {
        "title": "Brain",
        "default_activity": 100,
        "dilution": "2ml",
        "add Tc99m": "to vial B",
        "add 3ml saline": "to vial A",
        "shake": "few times",
        "add 1ml": "from A to B",
        "incubation": "30min",
        "final_volume": 3
    },
    "NEOSPECT": {
        "title": "Lungs SPECT",
        "default_activity": 30,
        "dilution": "1ml",
        "remove": "equal volume of air",
        "shake": "few times",
        "boil": "10min",
        "cooling": "15min",
        "final_volume": 1
    }
}
