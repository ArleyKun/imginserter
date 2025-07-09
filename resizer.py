import win32com.client as win32
import os
from tkinter import Tk, filedialog
from colorama import init, Fore, Style
from datetime import datetime

init(autoreset=True)

# sc
def log_action(message):
    with open("log.txt", "a") as log_file:
        log_file.write(f"[{datetime.now()}] {message}\n")

# made by arley
def show_banner():
    print(Fore.CYAN + r"""
                                                          _____                                                     
    _____        ___________     _____               _____\    \  ______   _____                                    
  /      |_      \          \   |\    \             /    / |    ||\     \ |     |                                   
 /         \      \    /\    \   \\    \           /    /  /___/|\ \     \|     |                                   
|     /\    \      |   \_\    |   \\    \         |    |__ |___|/ \ \           |                                   
|    |  |    \     |      ___/     \|    | ______ |       \        \ \____      |                                   
|     \/      \    |      \  ____   |    |/      \|     __/ __      \|___/     /|                                   
|\      /\     \  /     /\ \/    \  /            ||\    \  /  \         /     / |                                   
| \_____\ \_____\/_____/ |\______| /_____/\_____/|| \____\/    |       /_____/  /                                   
| |     | |     ||     | | |     ||      | |    ||| |    |____/|       |     | /                                    
 \|_____|\|_____||_____|/ \|_____||______|/|____|/ \|____|   | |       |_____|/                                     
                                                         |___|/                                                     
                         _____                                                              _____                   
___________         _____\    \            _____    ____________    _____  ______      _____\    \ ___________      
\          \       /    / |    |      _____\    \  /            \  /    / /     /|    /    / |    |\          \     
 \    /\    \     /    /  /___/|     /    / \    ||\___/\  \\___/||     |/     / |   /    /  /___/| \    /\    \    
  |   \_\    |   |    |__ |___|/    |    |  /___/| \|____\  \___|/|\____\\    / /   |    |__ |___|/  |   \_\    |   
  |      ___/    |       \       ____\    \ |   ||       |  |      \|___|/   / /    |       \        |      ___/    
  |      \  ____ |     __/ __   /    /\    \|___|/  __  /   / __      /     /_/____ |     __/ __     |      \  ____ 
 /     /\ \/    \|\    \  /  \ |    |/ \    \      /  \/   /_/  |    /     /\      ||\    \  /  \   /     /\ \/    \
/_____/ |\______|| \____\/    ||\____\ /____/|    |____________/|   /_____/ /_____/|| \____\/    | /_____/ |\______|
|     | | |     || |    |____/|| |   ||    | |    |           | /   |    |/|     | || |    |____/| |     | | |     |
|_____|/ \|_____| \|____|   | | \|___||____|/     |___________|/    |____| |_____|/  \|____|   | | |_____|/ \|_____| 
                        |___|/                                                             |___|/                    
""" + Style.RESET_ALL)


#inputs
def get_user_input():
    print(Fore.CYAN + Style.BRIGHT + "\nChoose Paper Size:")
    print(Fore.WHITE + "  1 - " + Fore.LIGHTGREEN_EX + "SHORT (8.5 x 11)")
    print(Fore.WHITE + "  2 - " + Fore.LIGHTGREEN_EX + "A4    (8.27 x 11.69)")
    print(Fore.WHITE + "  3 - " + Fore.LIGHTGREEN_EX + "LONG  (8.5 x 13)")
    paper = input(Fore.YELLOW + Style.BRIGHT + "\nEnter paper size (1/2/3): ").strip()

    print(Fore.CYAN + Style.BRIGHT + "\nChoose Image Size:")
    print(Fore.WHITE + "  1 - " + Fore.LIGHTMAGENTA_EX + "Full Size (1 per page)")
    print(Fore.WHITE + "  2 - " + Fore.LIGHTMAGENTA_EX + "Half Size (2 per page)")
    size = input(Fore.YELLOW + Style.BRIGHT + "\nEnter image size (1/2): ").strip()

    if paper not in ['1', '2', '3'] or size not in ['1', '2']:
        print(Fore.RED + Style.BRIGHT + "\nInvalid input. Exiting.")
        exit()

    return paper, size


#selct
def select_images():
    Tk().withdraw()
    paths = filedialog.askopenfilenames(title="Select Image(s)", filetypes=[("Image Files", "*.jpg *.jpeg *.png")])
    if not paths:
        print(Fore.RED + "No images selected. Exiting.")
        exit()
    return paths

#paper sizes
page_sizes = {
    '1': (8.5, 11),      # SHORT
    '2': (8.27, 11.69),  # A4
    '3': (8.5, 13)       # LONG
}

margins = {
    '1': {'top': 0.25, 'bottom': 0.5, 'left': 0.25, 'right': 0.5},
    '2': {'top': 0.25, 'bottom': 1.0, 'left': 0.19, 'right': 1.0},
    '3': {'top': 0.25, 'bottom': 0.5, 'left': 0.25, 'right': 0.5}
}

image_sizes = {
    '1': {'1': (8, 10.5), '2': (8, 5.25)},
    '2': {'1': (7.9, 11.13), '2': (7.9, 5.65)},
    '3': {'1': (8, 12.5), '2': (8, 6.25)}
}

#page set
def setup_page(doc, paper_code):
    section = doc.PageSetup
    width, height = page_sizes[paper_code]
    margin = margins[paper_code]

    section.PageWidth = width * 72
    section.PageHeight = height * 72
    section.TopMargin = margin['top'] * 72 #cv to pts
    section.BottomMargin = margin['bottom'] * 72
    section.LeftMargin = margin['left'] * 72
    section.RightMargin = margin['right'] * 72

#full size
def insert_full_size_images(doc, sel, image_paths, img_width, img_height):
    for i, img_path in enumerate(image_paths):
        sel.ParagraphFormat.SpaceBefore = 0
        sel.ParagraphFormat.SpaceAfter = 0
        sel.ParagraphFormat.LineSpacingRule = 0

        shape = sel.InlineShapes.AddPicture(
            FileName=os.path.abspath(img_path),
            LinkToFile=False,
            SaveWithDocument=True
        )
        shape.Width = int(img_width)
        shape.Height = int(img_height)
        sel.TypeParagraph()

        if i < len(image_paths) - 1:
            sel.InsertBreak(7)
            sel.Collapse(0)

#clean
    for p in doc.Paragraphs:
        if p.Range.Text.strip() == "":
            p.Range.Delete()

#half size
def insert_half_size_images(doc, image_paths, img_width, img_height, page_width, page_height):
    total = len(image_paths)
    i = 0

    img_width_pt = img_width * 72
    img_height_pt = img_height * 72
    page_width_pt = page_width * 72
    page_height_pt = page_height * 72

    left = (page_width_pt - img_width_pt) / 2
    top1 = (page_height_pt - (2 * img_height_pt)) / 3
    top2 = top1 * 2 + img_height_pt

    while i < total:
        doc.Shapes.AddPicture(
            FileName=os.path.abspath(image_paths[i]),
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top1,
            Width=img_width_pt,
            Height=img_height_pt
        )
        log_action(f"Inserted half-size image {os.path.basename(image_paths[i])} (top)")

        if i + 1 < total:
            doc.Shapes.AddPicture(
                FileName=os.path.abspath(image_paths[i + 1]),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=left,
                Top=top2,
                Width=img_width_pt,
                Height=img_height_pt
            )
            log_action(f"Inserted half-size image {os.path.basename(image_paths[i + 1])} (bottom)")

        i += 2
        if i < total:
            doc.Paragraphs.Add()
            doc.Range(doc.Content.End - 1).InsertBreak(7)

#run
def main():
    show_banner()
    paper_input, size_input = get_user_input()
    image_paths = select_images()

    word = win32.Dispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Add()
    sel = word.Selection

    setup_page(doc, paper_input)
    img_width_in, img_height_in = image_sizes[paper_input][size_input]

    if size_input == '1':
        insert_full_size_images(doc, sel, image_paths, img_width_in * 72, img_height_in * 72)
    else:
        page_width, page_height = page_sizes[paper_input]
        insert_half_size_images(doc, image_paths, img_width_in, img_height_in, page_width, page_height)

    print(Fore.CYAN + f"\nInserted {len(image_paths)} image(s).\n" + Style.RESET_ALL)

if __name__ == "__main__":
    main()
