from pptx import Presentation

def inspect_layouts(template_path):
    prs = Presentation(template_path)
    print(f"Inspecting '{template_path}'...")
    print(f"Number of slide layouts: {len(prs.slide_layouts)}\n")
    
    for i, layout in enumerate(prs.slide_layouts):
        print(f"--- Layout {i}: {layout.name} ---")
        for ph in layout.placeholders:
            print(f"  Placeholder index {ph.placeholder_format.idx}: name='{ph.name}', type={ph.placeholder_format.type}")
        print("")

if __name__ == "__main__":
    inspect_layouts("Chapter 8.pptx")
