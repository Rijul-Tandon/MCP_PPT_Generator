import win32com.client
import os

def test_com_replace():
    path = os.path.abspath('output/clone_test.pptx')
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    prs = ppt.Presentations.Open(path, WithWindow=False)
    
    slide = prs.Slides(2) # 1-based, 2nd slide
    replacements = {
        "Leveraging the referral network": "Using the new MCP system"
    }
    
    replaced = 0
    for shape in slide.Shapes:
        if shape.HasTextFrame:
            if shape.TextFrame.HasText:
                txt_range = shape.TextFrame.TextRange
                for old_t, new_t in replacements.items():
                    # Replace returns a TextRange object if found, or None
                    found = txt_range.Replace(FindWhat=old_t, ReplaceWhat=new_t, WholeWords=False)
                    if found is not None:
                        replaced += 1
                        print(f"Replaced in shape '{shape.Name}'")
    
    prs.Save()
    prs.Close()
    ppt.Quit()
    print("Total replaced:", replaced)

if __name__ == "__main__":
    test_com_replace()
