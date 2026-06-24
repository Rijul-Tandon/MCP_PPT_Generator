import win32com.client
import os

def clone_test():
    base_path = os.path.abspath('context/Referral_Analysis/reference_decks/2026-05-14 NT1 Referral Network Analysis - Insights & Recommendations v1.pptx')
    out_path = os.path.abspath('output/clone_test.pptx')
    
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    # ppt.Visible = True # Commented out to run silently if possible
    
    # Open presentation
    prs = ppt.Presentations.Open(base_path, WithWindow=False)
    
    # Let's say we want slides [0, 4, 4, 9] mapped to 1-based index: 1, 5, 5, 10
    wanted_indices = [1, 5, 5, 10]
    
    # The trick: it's easier to duplicate the ones we want to the end, then delete all the original slides!
    orig_count = prs.Slides.Count
    for idx in wanted_indices:
        # Duplicate to the end
        prs.Slides(idx).Duplicate().MoveTo(prs.Slides.Count)
        
    # Delete the original slides (from end to start so indices don't shift)
    for i in range(orig_count, 0, -1):
        prs.Slides(i).Delete()
        
    prs.SaveAs(out_path)
    prs.Close()
    ppt.Quit()
    print("Cloned successfully to", out_path)

if __name__ == '__main__':
    clone_test()
