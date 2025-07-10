import win32com.client as win32
import os
from tkinter import Tk, filedialog
from colorama import init, Fore, Style
from datetime import datetime

init(autoreset=True)

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

def get_user_input():
    print(Fore.GREEN + "Choose Paper Size:")
    print("1 - SHORT (8.5 x 11)")
    print("2 - A4    (8.27 x 11.69)")
    print("3 - LONG  (8.5 x 13)")
    paper = input(Fore.YELLOW + "Enter paper size (1/2/3): ").strip()

    print(Fore.GREEN + "\nChoose Layout:")
    print("1 - Full Size")
    print("2 - Half Size")
    print("3 - 4 Pics Layout (2x2 Grid)")
    size = input(Fore.YELLOW + "Enter layout option (1/2/3): ").strip()

    if paper not in ['1', '2', '3'] or size not in ['1', '2', '3']:
        print(Fore.RED + "Invalid input. Exiting.")
        exit()

    return paper, size

def select_images():
    Tk().withdraw()
    paths = filedialog.askopenfilenames(title="Select Image(s)", filetypes=[("Image Files", "*.jpg *.jpeg *.png")])
    if not paths:
        print(Fore.RED + "No images selected. Exiting.")
        exit()
    return paths

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

grid_image_sizes = {
    '1': (5.25, 3.94),  # SHORT
    '2': (5.57, 3.75),  # A4
    '3': (6.24, 3.95),  # LONG
}

def setup_page(doc, paper_code, landscape=False):
    section = doc.PageSetup
    width, height = page_sizes[paper_code]
    margin = margins[paper_code]

    if landscape:
        width, height = height, width

    section.PageWidth = width * 72
    section.PageHeight = height * 72
    section.TopMargin = margin['top'] * 72
    section.BottomMargin = margin['bottom'] * 72
    section.LeftMargin = margin['left'] * 72
    section.RightMargin = margin['right'] * 72


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

    for p in doc.Paragraphs:
        if p.Range.Text.strip() == "":
            p.Range.Delete()

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

# 2x2grid
def insert_grid_images(doc, image_paths, img_width, img_height, page_width, page_height):
    img_width_pt = img_width * 72
    img_height_pt = img_height * 72
    margin_left = 0.25 * 72
    margin_top = 0.19 * 72
    spacing_x = 0.2 * 72  # columns
    spacing_y = 0.2 * 72  # rows

    i = 0
    total = len(image_paths)

    while i < total:
        col1_left = margin_left
        col2_left = margin_left + img_width_pt + spacing_x
        row1_top = margin_top
        row2_top = margin_top + img_height_pt + spacing_y

        # Tleft
        if i < total:
            doc.Shapes.AddPicture(
                FileName=os.path.abspath(image_paths[i]),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=col1_left,
                Top=row1_top,
                Width=img_width_pt,
                Height=img_height_pt
            )
            log_action(f"Inserted top-left: {os.path.basename(image_paths[i])}")

        # Tright
        if i + 1 < total:
            doc.Shapes.AddPicture(
                FileName=os.path.abspath(image_paths[i + 1]),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=col2_left,
                Top=row1_top,
                Width=img_width_pt,
                Height=img_height_pt
            )
            log_action(f"Inserted top-right: {os.path.basename(image_paths[i + 1])}")

        # Bleft
        if i + 2 < total:
            doc.Shapes.AddPicture(
                FileName=os.path.abspath(image_paths[i + 2]),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=col1_left,
                Top=row2_top,
                Width=img_width_pt,
                Height=img_height_pt
            )
            log_action(f"Inserted bottom-left: {os.path.basename(image_paths[i + 2])}")

        # Bright
        if i + 3 < total:
            doc.Shapes.AddPicture(
                FileName=os.path.abspath(image_paths[i + 3]),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=col2_left,
                Top=row2_top,
                Width=img_width_pt,
                Height=img_height_pt
            )
            log_action(f"Inserted bottom-right: {os.path.basename(image_paths[i + 3])}")

        i += 4
        if i < total:
            doc.Paragraphs.Add()
            doc.Range(doc.Content.End - 1).InsertBreak(7)

def main():
    show_banner()
    paper_code, size_code = get_user_input()
    image_paths = select_images()

    word = win32.gencache.EnsureDispatch("Word.Application")
    doc = word.Documents.Add()
    sel = word.Selection

    # ls opt3
    is_landscape = size_code == '3'
    setup_page(doc, paper_code, landscape=is_landscape)

    width, height = page_sizes[paper_code]
    if is_landscape:
        width, height = height, width # word calc

    if size_code == '1':  # full size
        img_width, img_height = image_sizes[paper_code]['1']
        insert_full_size_images(doc, sel, image_paths, img_width * 72, img_height * 72)

    elif size_code == '2':  # half size
        img_width, img_height = image_sizes[paper_code]['2']
        insert_half_size_images(doc, image_paths, img_width, img_height, width, height)

    elif size_code == '3':  # 4 pics Layout
        img_width, img_height = grid_image_sizes[paper_code]
        insert_grid_images(doc, image_paths, img_width, img_height, width, height)


    word.Visible = True
    log_action("Finished successfully.")

if __name__ == "__main__":
    main()
