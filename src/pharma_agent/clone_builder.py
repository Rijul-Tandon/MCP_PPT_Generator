"""
clone_builder.py - Win32COM-based slide cloning for 100% design fidelity.
"""

import os
from pathlib import Path
import win32com.client

def clone_deck(reference_path: str, output_path: str, slide_mapping: list[int]) -> str:
    """
    Creates a new presentation by perfectly duplicating the chosen slides from 
    a reference deck using PowerPoint COM automation. This guarantees 100% 
    fidelity of layouts, theme colors, charts, and relationships.
    
    reference_path: Absolute path to the source .pptx
    output_path: Absolute path to save the new .pptx
    slide_mapping: List of 0-based slide indices to keep, in the desired order.
                   e.g., [0, 4, 4, 9] creates a 4-slide deck.
    """
    ref_path = os.path.abspath(reference_path)
    out_path = os.path.abspath(output_path)
    
    if not os.path.exists(ref_path):
        raise FileNotFoundError(f"Reference deck not found: {ref_path}")
        
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        raise RuntimeError(f"Failed to connect to PowerPoint via COM. Is PowerPoint installed? Error: {e}")
        
    try:
        # Open presentation silently
        prs = ppt.Presentations.Open(ref_path, WithWindow=False)
        orig_count = prs.Slides.Count
        
        # Win32COM collections are 1-based.
        # wanted_indices maps our 0-based python indices to 1-based PPT indices
        wanted_indices = [i + 1 for i in slide_mapping]
        
        for idx in wanted_indices:
            if idx < 1 or idx > orig_count:
                raise ValueError(f"Slide index {idx-1} is out of bounds (deck has {orig_count} slides).")
                
            # Duplicate the slide to the very end of the presentation
            prs.Slides(idx).Duplicate().MoveTo(prs.Slides.Count)
            
        # Delete all the original slides (iterate backwards so indices don't shift)
        for i in range(orig_count, 0, -1):
            prs.Slides(i).Delete()
            
        prs.SaveAs(out_path)
        prs.Close()
        return f"Successfully cloned deck to {out_path} with {len(slide_mapping)} slides."
        
    finally:
        try:
            ppt.Quit()
        except:
            pass
