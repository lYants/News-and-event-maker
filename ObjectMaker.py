import yaml, json, os, sys, requests, subprocess, shutil, tkinter as tk, customtkinter as ctk
from appdirs import user_config_dir
from calendar import Calendar
from ctypes import windll
from datetime import datetime
from pathlib import Path
from tkcalendar import Calendar
from tkinter import ttk, filedialog

VERSION = "v1.0.2"
RELEASE_URL = "https://api.github.com/repos/lYants/News-and-event-maker/releases/latest"

DEFAULT_FONT = ("Sabon", 18)
OBJECT_CHOICES = ("news", "event")
LANGUAGES = ("nl", "en")
CATEGORIES = (
    "agriculture",
    "algorithms",
    "art",
    "AI",
    "care",
    "chatbot",
    "computational thinking",
    "kiks",
    "math with python",
    "physical computing",
    "python programming",
    "social robot",
    "stem",
    "wegostem",
)
CATEGORY_MAPPER = {
    "agriculture": "/images/curricula/logo_agriculture.png",
    "algorithms": "/images/curricula/logo_algorithms.png",
    "art": "/images/curricula/logo_art.png",
    "AI": "/images/curricula/logo_basics_ai.png",
    "care": "/images/curricula/logo_care.png",
    "chatbot": "/images/curricula/logo_chatbot.png",
    "computational thinking": "/images/curricula/logo_computational_thinking.png",
    "kiks": "/images/curricula/logo_kiks.png",
    "math with python": "/images/curricula/logo_math_with_python.png",
    "physical computing": "/images/curricula/logo_physical_computing.png",
    "python programming": "/images/curricula/logo_python_programming.png",
    "social robot": "/images/curricula/logo_socialrobot.png",
    "stem": "/images/curricula/logo_stem.png",
    "wegostem": "/images/curricula/logo_wegostem.png",
}

START_DATE, END_DATE = 0, 1
NEWS_OBJECT, EVENT_OBJECT = "news", "event"

base_path = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)


class WebObject:
    def __init__(self, title, language, anchor, category, date, year, content):
        (
            self.title,
            self.language,
            self.anchor,
            self.category,
            self.date,
            self.year,
            self.content,
        ) = (title, language, anchor, category, date, year, content)
        self.type = "news"

    def makeDict(self):
        return {
            "title": self.title,
            "language": self.language,
            "item_theme_logo_ur": self.category,
            "anchor": self.anchor,
            "date": self.date,
        }


class Event(WebObject):
    def __init__(
        self,
        title,
        language,
        anchor,
        category,
        date,
        enddate,
        location,
        loclink,
        reglink,
        year,
        month,
        content,
    ):
        super().__init__(title, language, anchor, category, date, year, content)
        self.type = "event"
        (
            self.endDate,
            self.location,
            self.locationLink,
            self.registrationLink,
            self.month,
        ) = (enddate, location, loclink, reglink, month)

    def makeDict(self):
        dic = super().makeDict()
        dic.update(
            {
                "end_date": self.endDate,
                "location": self.location,
                "location_link": self.locationLink,
                "registration_link": self.registrationLink,
            }
        )
        return dic


class Gui:
    def __init__(self, objectmaker) -> None:
        self.root = tk.Tk()
        self.root.tk.call(
            "source",
            os.path.join(base_path, "styles", "forest-light.tcl"),
        )
        self.style = ttk.Style()
        self.style.theme_use("forest-light")
        self.root.title("Object maker")
        self.root.geometry("1800x775")
        windll.shcore.SetProcessDpiAwareness(1)
        self.root.option_add("*TCombobox*Listbox.font", DEFAULT_FONT)
        self.variableComponents = []
        self.warning = None
        self.month = None
        self.year = None

        self.mainframe = ttk.Frame(self.root, padding=(3, 3, 12, 12))
        self.mainframe.grid(column=0, row=0, sticky="NWSE")

        sep = ttk.Separator(self.mainframe, orient="vertical")
        sep.grid(row=0, column=4, rowspan=12, sticky="ns", padx=20)

        self.style.configure(".", font=("Sabon", 18))

        self.maker = objectmaker
        self.makeMenuBar()
        self.makeBasicComponents()
        self.makeTextEditor()
        self.checkWebsiteDir()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.mainloop()

    def close(self):
        self.maker.saveSettings()
        self.root.destroy()

    def makeMenuBar(self):
        self.menu = tk.Menu(self.root)
        self.menu.add_command(
            label="Select Dwengo-website folder", command=self.getWebsiteDir
        )
        self.menu.add_command(
            label=f"\tcurrent: {self.maker.websiteDir}", state="disabled"
        )
        self.siteIndex = self.menu.index("end")
        self.root.config(menu=self.menu)

    def getWebsiteDir(self):
        if self.siteIndex is None:
            return
        self.maker.selectWebsiteDir()
        self.menu.entryconfig(
            self.siteIndex, label=f"\tcurrent: {self.maker.websiteDir}"
        )
        self.checkWebsiteDir()

    def makeBasicComponents(self):
        # Object selection
        self.objectType = tk.StringVar()
        self.objectBox = ttk.Combobox(
            self.mainframe,
            textvariable=self.objectType,
            values=OBJECT_CHOICES,
            font=DEFAULT_FONT,
            state="readonly",
            width=15,
        )
        self.objectBox.current(0)
        self.objectBox.grid(column=0, row=0, padx=5, pady=5)
        self.objectBox.bind("<<ComboboxSelected>>", self.removeHighlight)
        self.setNewsComponents()
        self.currentObject = NEWS_OBJECT

        # Title input
        ttk.Label(self.mainframe, text="Title:", font=DEFAULT_FONT).grid(
            column=0, row=1, sticky="W", padx=5, pady=5
        )
        self.title = ttk.Entry(self.mainframe, font=DEFAULT_FONT, width=30)
        self.title.grid(column=1, row=1, sticky="ew", padx=5, pady=5, columnspan=2)

        # Category selection
        ttk.Label(self.mainframe, text="Category:", font=DEFAULT_FONT).grid(
            column=0, row=2, sticky="W", padx=5, pady=5
        )
        self.category = tk.StringVar()
        box = ttk.Combobox(
            self.mainframe,
            textvariable=self.category,
            values=CATEGORIES,
            font=DEFAULT_FONT,
            state="readonly",
        )
        box.grid(column=1, row=2, padx=5, pady=5, columnspan=2, sticky="eW")
        box.bind("<<ComboboxSelected>>", self.removeHighlight)

        # Language selection
        ttk.Label(self.mainframe, text=f"Language:", font=DEFAULT_FONT).grid(
            column=0, row=3, sticky="W", padx=5, pady=5
        )
        self.language = tk.StringVar()
        languageBox = ttk.Combobox(
            self.mainframe,
            textvariable=self.language,
            values=LANGUAGES,
            font=DEFAULT_FONT,
            state="readonly",
            width=3,
        )
        languageBox.current(0)
        languageBox.grid(column=1, row=3, padx=5, pady=5, columnspan=2, sticky="W")
        languageBox.bind("<<ComboboxSelected>>", self.removeHighlight)

        # Anchor input
        ttk.Label(self.mainframe, text="Anchor:", font=DEFAULT_FONT).grid(
            column=0, row=4, sticky="W", padx=5, pady=5
        )
        self.anchor = ttk.Entry(self.mainframe, font=DEFAULT_FONT)
        self.anchor.grid(column=1, row=4, padx=5, pady=5, columnspan=2, sticky="eW")

    def makeTextEditor(self):
        self.textEditor = tk.Text(self.mainframe, font=DEFAULT_FONT, undo=True)
        self.textEditor.bind("<Control-z>", lambda event: self.textEditor.edit_undo())
        self.textEditor.bind("<Control-y>", lambda event: self.textEditor.edit_redo())
        self.textEditor.grid(column=5, row=1, rowspan=11, columnspan=5)
        frame = ttk.Frame(self.mainframe)
        frame.grid(column=5, row=0, columnspan=5, sticky="news", padx=20)
        ctk.CTkButton(
            frame,
            text="bold",
            fg_color="black",
            text_color="white",
            hover_color="#222222",
            font=("Sabon", 14, "bold"),
            corner_radius=10,
            width=15,
            height=10,
            command=self.makeTextBold,
        ).pack(side="left", padx=10)
        ctk.CTkButton(
            frame,
            text="italic",
            fg_color="black",
            text_color="white",
            hover_color="#222222",
            font=("Sabon", 14, "italic"),
            corner_radius=10,
            width=15,
            height=10,
            command=self.makeTextItalic,
        ).pack(side="left", padx=10)
        ctk.CTkButton(
            frame,
            text="insert image",
            fg_color="black",
            text_color="white",
            hover_color="#222222",
            font=("Sabon", 14),
            corner_radius=10,
            width=15,
            height=10,
            command=self.insertImage,
        ).pack(side="left", padx=10)
        ctk.CTkButton(
            frame,
            text="insert weblink",
            fg_color="black",
            text_color="white",
            hover_color="#222222",
            font=("Sabon", 14),
            corner_radius=10,
            width=15,
            height=10,
            command=self.insertLink,
        ).pack(side="left", padx=10)
        ctk.CTkButton(
            frame,
            text="insert break",
            fg_color="black",
            text_color="white",
            hover_color="#222222",
            font=("Sabon", 14),
            corner_radius=10,
            width=15,
            height=10,
            command=self.insertBreak,
        ).pack(side="left", padx=10)

    def updateObjectType(self):
        value = self.objectType.get()

        if value == NEWS_OBJECT and self.currentObject != NEWS_OBJECT:
            for i in self.variableComponents:
                i.destroy()

            self.variableComponents = []
            self.currentObject = NEWS_OBJECT
            self.setNewsComponents()

        elif value == EVENT_OBJECT and self.currentObject != EVENT_OBJECT:
            for i in self.variableComponents:
                i.destroy()

            self.variableComponents = []
            self.currentObject = EVENT_OBJECT
            self.setEventComponents()

    def setNewsComponents(self):
        # Date
        dateLabel = ttk.Label(self.mainframe, text=f"Date:", font=DEFAULT_FONT)
        dateLabel.grid(column=0, row=5, sticky="W", padx=5, pady=5)
        self.variableComponents.append(dateLabel)

        date = datetime.now()
        self.year = date.year
        self.formattedDate = date.strftime("%Y-%m-%dT12:00:00")

        dateEntry = ttk.Label(
            self.mainframe, text=f"{self.formattedDate}", font=DEFAULT_FONT
        )
        dateEntry.grid(column=1, row=5, sticky="eW", padx=5, pady=5, columnspan=2)
        self.variableComponents.append(dateEntry)

        # Make object button
        makeButton = ttk.Button(
            self.mainframe, text="make object", command=self.makeNewsObject
        )
        makeButton.grid(column=2, row=0, padx=5, pady=5, sticky="e")
        self.variableComponents.append(makeButton)

    def setEventComponents(self):
        # Start date
        self.formattedDate = ""
        startDateLabel = ttk.Label(
            self.mainframe, text=f"Start date:", font=DEFAULT_FONT
        )
        startDateLabel.grid(column=0, row=5, sticky="W", padx=5, pady=5)
        self.variableComponents.append(startDateLabel)

        startDateButton = ttk.Button(
            self.mainframe,
            text="select",
            command=lambda: self.askDate(START_DATE),
            width=8,
        )
        startDateButton.grid(column=2, row=5, sticky="e", padx=5, pady=5)
        self.variableComponents.append(startDateButton)

        # Start time
        startTimeLabel = ttk.Label(
            self.mainframe, text="Start time:", font=DEFAULT_FONT
        )
        startTimeLabel.grid(column=0, row=6, sticky="W", padx=5, pady=5)
        self.variableComponents.append(startTimeLabel)

        startTimeFrame = ttk.Frame(self.mainframe)
        startTimeFrame.grid(row=6, column=1, sticky="W", padx=5, pady=10)
        self.variableComponents.append(startTimeFrame)

        self.startTimeH = ttk.Spinbox(
            startTimeFrame, from_=0, to=23, increment=1, font=DEFAULT_FONT, width=3
        )
        self.startTimeH.pack(side="left")
        self.startTimeH.set(12)

        ttk.Label(startTimeFrame, text=":", font=DEFAULT_FONT).pack(side="left")

        self.startTimeM = ttk.Spinbox(
            startTimeFrame, from_=0, to=45, increment=15, font=DEFAULT_FONT, width=3
        )
        self.startTimeM.pack(side="right")
        self.startTimeM.set(00)

        # End date
        self.formattedEndDate = ""
        endDateLabel = ttk.Label(self.mainframe, text=f"Date:", font=DEFAULT_FONT)
        endDateLabel.grid(column=0, row=7, sticky="W", padx=5, pady=5)
        self.variableComponents.append(endDateLabel)

        endDateButton = ttk.Button(
            self.mainframe,
            text="select",
            command=lambda: self.askDate(END_DATE),
            width=8,
        )
        endDateButton.grid(column=2, row=7, sticky="e", padx=5, pady=5)
        self.variableComponents.append(endDateButton)

        # End time
        endTimeLabel = ttk.Label(self.mainframe, text="End time:", font=DEFAULT_FONT)
        endTimeLabel.grid(column=0, row=8, sticky="W", padx=5, pady=5)
        self.variableComponents.append(endTimeLabel)

        endTimeFrame = ttk.Frame(self.mainframe)
        endTimeFrame.grid(row=8, column=1, sticky="W", padx=5, pady=10)
        self.variableComponents.append(endTimeFrame)

        self.endTimeH = ttk.Spinbox(
            endTimeFrame, from_=0, to=23, increment=1, font=DEFAULT_FONT, width=3
        )
        self.endTimeH.pack(side="left")
        self.endTimeH.set(12)

        ttk.Label(endTimeFrame, text=":", font=DEFAULT_FONT).pack(side="left")

        self.endTimeM = ttk.Spinbox(
            endTimeFrame, from_=0, to=45, increment=15, font=DEFAULT_FONT, width=3
        )
        self.endTimeM.pack(side="right")
        self.endTimeM.set(00)

        # Location
        locationLabel = ttk.Label(self.mainframe, text="Location:", font=DEFAULT_FONT)
        locationLabel.grid(column=0, row=9, sticky="W", padx=5, pady=5)
        self.variableComponents.append(locationLabel)

        self.location = ttk.Entry(self.mainframe, font=DEFAULT_FONT)
        self.location.grid(column=1, row=9, sticky="ew", padx=5, pady=5, columnspan=2)
        self.variableComponents.append(self.location)

        # Location link
        locationLinkLabel = ttk.Label(
            self.mainframe, text="Location link:", font=DEFAULT_FONT
        )
        locationLinkLabel.grid(column=0, row=10, sticky="W", padx=5, pady=5)
        self.variableComponents.append(locationLinkLabel)

        self.locationLink = ttk.Entry(self.mainframe, font=DEFAULT_FONT)
        self.locationLink.grid(
            column=1, row=10, sticky="ew", padx=5, pady=5, columnspan=2
        )
        self.variableComponents.append(self.locationLink)

        # Registration link
        registrationLabel = ttk.Label(
            self.mainframe, text="Registration link:", font=DEFAULT_FONT
        )
        registrationLabel.grid(column=0, row=11, sticky="W", padx=5, pady=5)
        self.variableComponents.append(registrationLabel)

        self.registrationLink = ttk.Entry(
            self.mainframe,
            font=DEFAULT_FONT,
        )
        self.registrationLink.grid(
            column=1, row=11, sticky="ew", padx=5, pady=5, columnspan=2
        )
        self.variableComponents.append(self.registrationLink)

        # Make object button
        makeButton = ttk.Button(
            self.mainframe, text="make object", command=self.makeEventObject
        )
        makeButton.grid(column=2, row=0, padx=5, pady=5, sticky="e")
        self.variableComponents.append(makeButton)

    def makeNewsObject(self):
        category = self.category.get()
        if category:
            image = CATEGORY_MAPPER[category]
        else:
            image = ""

        object = WebObject(
            title=self.title.get(),
            date=self.formattedDate,
            language=self.language.get(),
            anchor=self.anchor.get(),
            category=image,
            year=self.year,
            content=self.textEditor.get("1.0", "end-1c"),
        )
        self.maker.setObject(object)
        self.askFileName()

    def makeEventObject(self):
        category = self.category.get()
        if category:
            image = CATEGORY_MAPPER[self.category.get()]
        else:
            image = ""

        if not self.month:
            self.month = datetime.now().month
        if not self.year:
            self.year = datetime.now().year

        startHour = self.startTimeH.get()
        startMinute = self.startTimeM.get()
        endHour = self.endTimeH.get()
        endMinute = self.endTimeM.get()

        if len(startHour) == 1:
            startHour = '0' + startHour
        if len(startMinute) == 1:
            startMinute = '0' + startMinute
        if len(endHour) == 1:
            endHour = '0' + endHour
        if len(endMinute) == 1:
            endMinute = '0' + endMinute
        object = Event(
            title=self.title.get(),
            date=self.formattedDate
            + "T"
            + startHour
            + ":"
            + startMinute,
            enddate=self.formattedEndDate
            + "T"
            + endHour
            + ":"
            + endMinute,
            language=self.language.get(),
            anchor=self.anchor.get(),
            category=image,
            location=self.location.get(),
            loclink=self.locationLink.get(),
            reglink=self.registrationLink.get(),
            year=self.year,
            month=self.month,
            content=self.textEditor.get("1.0", "end-1c"),
        )
        self.maker.setObject(object)
        self.askFileName()

    def askFileName(self):
        if self.warning is not None:
            return

        root = tk.Toplevel()

        ttk.Label(root, text="Filename:", font=DEFAULT_FONT).grid(
            column=0, row=0, sticky="W", padx=5, pady=5
        )
        filename = ttk.Entry(root, font=DEFAULT_FONT)
        filename.grid(column=1, row=0, sticky="ew", padx=5, pady=5)

        ttk.Button(
            root,
            text="make file",
            command=lambda: self.getFilenameAndMakeFile(filename.get()),
        ).grid(column=2, row=0, padx=5, pady=5, sticky="e")
        root.mainloop()

    def checkWebsiteDir(self):
        if self.maker.websiteDir is None:
            self.warning = tk.Label(
                self.mainframe,
                text="No dwengo-website folder selected!",
                font=("Sabon", 15),
                fg="red",
            )
            self.warning.grid(column=0, row=12, sticky="W", padx=5, pady=5)
        else:
            if self.warning:
                self.warning.destroy()
                self.warning = None

    def getFilenameAndMakeFile(self, filename):
        self.maker.makeObjectFile(filename)
        self.close()

    def askDate(self, type):
        self.calRoot = tk.Tk()
        self.calRoot.geometry("300x300")
        date = datetime.now()
        self.cal = Calendar(
            self.calRoot,
            selectmode="day",
            year=date.year,
            month=date.month,
            day=date.day,
        )
        self.cal.pack(pady=20)

        ttk.Button(
            self.calRoot, text="select", command=lambda: self.getDate(type)
        ).pack()
        self.calRoot.wait_window()

    def getDate(self, type):
        d = datetime.strptime(self.cal.get_date(), "%m/%d/%y")
        if type == START_DATE:
            self.month = d.month
            self.year = d.year
            self.formattedDate = d.strftime("%Y-%m-%d")
            ttk.Label(
                self.mainframe, text=f"{self.formattedDate}", font=DEFAULT_FONT
            ).grid(column=1, row=5, sticky="W", padx=5, pady=5)
        elif type == END_DATE:
            self.formattedEndDate = d.strftime("%Y-%m-%d")
            ttk.Label(
                self.mainframe, text=f"{self.formattedEndDate}", font=DEFAULT_FONT
            ).grid(column=1, row=7, sticky="W", padx=5, pady=5)
        self.calRoot.destroy()

    def makeTextBold(self):
        try:
            start = self.textEditor.index("sel.first")
            end = self.textEditor.index("sel.last")
            selected = self.textEditor.get(start, end)

            if selected.startswith("**") and selected.endswith("**"):
                newBold = selected[2:-2]
            else:
                newBold = f"**{selected}**"

            self.textEditor.delete(start, end)
            self.textEditor.insert(start, newBold)
        except tk.TclError:
            return

    def makeTextItalic(self):
        try:
            start = self.textEditor.index("sel.first")
            end = self.textEditor.index("sel.last")
            selected = self.textEditor.get(start, end)

            if selected.startswith("*") and selected.endswith("*"):
                newIt = selected[2:-2]
            else:
                newIt = f"*{selected}*"

            self.textEditor.delete(start, end)
            self.textEditor.insert(start, newIt)
        except tk.TclError:
            return

    def insertImage(self):
        self.textEditor.insert("insert", "![ image_description ]( /images/ )")

    def insertLink(self):
        self.textEditor.insert("insert", "[ link_description ]( link )")

    def insertBreak(self):
        self.textEditor.insert("insert","<br>")

    def removeHighlight(self, event):
        event.widget.selection_clear()
        self.mainframe.focus()
        self.updateObjectType()


class ObjectMaker:
    def __init__(self) -> None:
        self.loadSettings()

    def makeObjectFile(self, filename):
        if self.websiteDir is None or self.eventDir is None or self.newsDir is None:
            return

        if self.object.type == "event":
            if self.object.month >= 9:
                p = os.path.join(
                    self.eventDir, f"{self.object.year}" + "_najaar", filename + ".md"
                )
            else:
                p = os.path.join(
                    self.eventDir, f"{self.object.year}" + "_voorjaar", filename + ".md"
                )
        elif self.object.type == "news":
            p = os.path.join(self.newsDir, f"{self.object.year}", filename + ".md")

        preamble = self.object.makeDict()
        preambleYaml = yaml.dump(preamble)

        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as file:
            file.write("---\n" + preambleYaml + "---\n\n" + self.object.content)
        self.saveSettings()

    def loadSettings(self):
        configDir = user_config_dir("ObjectMaker", "Dwengo")
        os.makedirs(configDir, exist_ok=True)

        self.settingsFile = os.path.join(configDir, "settings.json")

        if not os.path.exists(self.settingsFile):
            self.setDefaultSettings()
        else:
            with open(self.settingsFile, "r") as f:
                settings = json.load(f)
                dir = settings["websiteDir"]
                if dir is None or not os.path.exists(dir):
                    self.setDefaultSettings()
                else:
                    self.websiteDir = dir
                    self.eventDir = settings["eventDir"]
                    self.newsDir = settings["newsDir"]

    def setDefaultSettings(self):
        self.websiteDir = None
        self.eventDir = None
        self.newsDir = None

    def saveSettings(self):
        with open(self.settingsFile, "w") as f:
            json.dump(
                {
                    "websiteDir": self.websiteDir,
                    "eventDir": self.eventDir,
                    "newsDir": self.newsDir,
                },
                f,
                indent=4,
            )

    def setObject(self, object):
        self.object = object

    def selectWebsiteDir(self):
        dir = filedialog.askdirectory(mustexist=True)
        if dir != "":
            self.websiteDir = dir
            self.eventDir = os.path.join(dir, "_events")
            self.newsDir = os.path.join(dir, "_news")
        self.saveSettings()


def newVersionAvailable(data):
    latest_version = data["tag_name"]

    if latest_version != VERSION:
        print("Nieuwe update beschikbaar!")
        return True
    return False


def updateExe(data):
    try:
        assets = {a["name"]: a["browser_download_url"] for a in data["assets"]}

        exe_dir = os.path.dirname(sys.executable)
        update_dir = os.path.join(exe_dir, "update_temp")
        os.makedirs(update_dir, exist_ok=True)

        for filename, url in assets.items():
            r = requests.get(url, stream=True)
            r.raise_for_status()
            with open(os.path.join(update_dir, filename), "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)

        shutil.move(
            os.path.join(update_dir, "updater.exe"),
            os.path.join(exe_dir, "updater.exe")
        )

    except Exception as e:
        print(f"Update download mislukt: {e}")
        return False
    return True


def runUpdater():
    exe_dir = os.path.dirname(sys.executable)
    current_exe = sys.executable
    new_exe_path = os.path.join(exe_dir, "update_temp", "objectmaker.exe")
    updater_path = os.path.join(exe_dir, "updater.exe")

    subprocess.Popen([updater_path, current_exe, new_exe_path])
    sys.exit()


if __name__ == "__main__":
    response = requests.get(RELEASE_URL)
    data = response.json()
    if newVersionAvailable(data):
        if updateExe(data):
            runUpdater()
        else:
            print("Update failed, continuing with current version.")

    objectmaker = ObjectMaker()
    obj = Gui(objectmaker)
