#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import glob
import re
import tempfile # For temporary files
import json
import subprocess
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
import mermaid # For generating SVG from Mermaid code
import cairosvg # For converting SVG to PNG
from datetime import date # For today's date

def convert_md_to_preview_html(md_content, theme="dark"):
    """
    Converts markdown content into HTML with preview styling.
    Supports 'dark' (VS Code Dark) and 'light' (GitHub Light).
    """
    lines = md_content.splitlines()
    title = ""
    meta_items = []
    mermaid_lines = []
    in_mermaid = False
    body_elements = []
    
    in_list = False
    in_sublist = False

    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("<!--"):
            continue
        if stripped == "---":
            continue

        if stripped == "```mermaid":
            in_mermaid = True
            continue
        if in_mermaid:
            if stripped == "```":
                in_mermaid = False
            else:
                mermaid_lines.append(line)
            continue

        if stripped.startswith("# "):
            title = stripped[2:].strip()
            continue

        if stripped.startswith("- 組織:") or stripped.startswith("- 職位:"):
            meta_items.append(stripped.lstrip("- ").strip())
            continue

        if stripped.startswith("Generate Date:"):
            if in_sublist:
                body_elements.append("</ul></li>")
                in_sublist = False
            if in_list:
                body_elements.append("</ul>")
                in_list = False
            body_elements.append(f"<div class=\"date-info\">{stripped}</div>")
            continue

        if stripped.startswith("### "):
            if in_sublist:
                body_elements.append("</ul></li>")
                in_sublist = False
            if in_list:
                body_elements.append("</ul>")
                in_list = False
            body_elements.append(f"<h3>{stripped[4:].strip()}</h3>")
            continue

        # Sub-list item (indented)
        if re.match(r"^\s+[-*]\s+", line):
            item_text = re.sub(r"^\s+[-*]\s+", "", line).strip()
            item_text = re.sub(r"\*\*(.*?)\*\*", r"<strong>\1</strong>", item_text)
            if not in_sublist:
                body_elements.append("<ul>")
                in_sublist = True
            body_elements.append(f"<li>{item_text}</li>")
            continue

        # Main list item
        if stripped.startswith("- ") or stripped.startswith("* "):
            if in_sublist:
                body_elements.append("</ul></li>")
                in_sublist = False
            if not in_list:
                body_elements.append("<ul>")
                in_list = True

            item_text = stripped[2:].strip()
            link_match = re.search(r"\[(.*?)\]\(.*?\)", item_text)
            if link_match:
                item_text = re.sub(r"\[(.*?)\]\(.*?\)", r"<a href=\"#\">\1</a>", item_text)
                body_elements.append(f"<li>{item_text}")
            else:
                body_elements.append(f"<li>{item_text}</li>")
            continue

    if in_sublist:
        body_elements.append("</ul></li>")
    if in_list:
        body_elements.append("</ul>")

    meta_html = "".join(f"<li>{item}</li>" for item in meta_items)
    mermaid_block = "\n".join(mermaid_lines)
    body_html = "\n  ".join(body_elements)

    if theme == "light":
        bg_color = "#ffffff"
        text_color = "#1f2328"
        h_color = "#1f2328"
        hr_color = "#d0d7de"
        a_color = "#0969da"
        date_color = "#656d76"
        mermaid_theme = "default"
    else:
        bg_color = "#1e1e1e"
        text_color = "#d4d4d4"
        h_color = "#ffffff"
        hr_color = "#3c3c3c"
        a_color = "#4ea1f3"
        date_color = "#858585"
        mermaid_theme = "dark"

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body {{
    background-color: {bg_color};
    color: {text_color};
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans", Roboto, "Microsoft YaHei", sans-serif;
    margin: 0;
    padding: 36px 48px;
    box-sizing: border-box;
    width: 1920px;
    min-height: 1080px;
  }}
  h1 {{
    font-size: 28px;
    font-weight: 600;
    color: {h_color};
    margin: 0 0 12px 0;
  }}
  hr {{
    border: none;
    border-top: 1px solid {hr_color};
    margin: 16px 0 24px 0;
  }}
  ul {{
    margin: 8px 0;
    padding-left: 24px;
    line-height: 1.6;
  }}
  li {{
    margin-bottom: 4px;
    font-size: 15px;
  }}
  h3 {{
    font-size: 20px;
    font-weight: 600;
    color: {h_color};
    margin: 24px 0 12px 0;
  }}
  a {{
    color: {a_color};
    text-decoration: none;
  }}
  .mermaid {{
    background-color: transparent;
    display: flex;
    justify-content: center;
    margin: 20px 0;
  }}
  .date-info {{
    color: {date_color};
    font-size: 14px;
    margin: 16px 0;
  }}
</style>
<script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
<script>
  mermaid.initialize({{
    startOnLoad: true,
    theme: "{mermaid_theme}",
    gantt: {{
      useWidth: 1600,
      fontSize: 14
    }}
  }});
</script>
</head>
<body>
  <h1>{title}</h1>
  <ul>
    {meta_html}
  </ul>
  <hr>
  <div class="mermaid">
{mermaid_block}
  </div>
  {body_html}
</body>
</html>
"""

def render_preview_image(md_content, image_path, theme="dark"):
    """
    Renders markdown content into a preview image using headless Chrome.
    Theme can be 'dark' (VS Code Dark) or 'light' (GitHub Light).
    """
    chrome_bin = (
        shutil.which("google-chrome")
        or shutil.which("chromium")
        or shutil.which("chromium-browser")
    )
    if not chrome_bin:
        print(f"Warning: Chrome/Chromium not found. Skipping {theme} preview image.")
        return False

    html_content = convert_md_to_preview_html(md_content, theme=theme)
    temp_html_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8") as temp_html:
            temp_html_path = temp_html.name
            temp_html.write(html_content)

        cmd = [
            chrome_bin,
            "--headless",
            "--no-sandbox",
            "--disable-gpu",
            "--virtual-time-budget=3000",
            "--window-size=1920,1080",
            f"--screenshot={image_path}",
            temp_html_path
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return os.path.exists(image_path)
    except Exception as e:
        print(f"Warning: Failed to render {theme} preview image: {e}")
        return False
    finally:
        if temp_html_path and os.path.exists(temp_html_path):
            os.remove(temp_html_path)

def create_single_slide_presentation(md_content, output_path):
    """
    Parses markdown content to create a single-slide presentation.
    The slide will contain the Gantt chart as an image at the top and all other
    relevant information below, with a font size of 12.
    """
    lines = md_content.split('\n')
    
    title = "Work Status Report" # Default title
    mermaid_code = ""
    other_content_lines = []

    # --- 1. Parse Markdown Content ---

    # Find the main title (H1)
    for line in lines:
        if line.startswith('# '):
            title = line[2:].strip() # This already removes '# '
            break
            
    # Extract the content of the Mermaid block
    in_mermaid = False
    mermaid_block_start_line = -1
    mermaid_block_end_line = -1
    for i, line in enumerate(lines):
        if line.strip() == '```mermaid':
            in_mermaid = True
            mermaid_block_start_line = i
            continue
        if in_mermaid and line.strip() == '```':
            in_mermaid = False
            mermaid_block_end_line = i
            break
        if in_mermaid:
            mermaid_code += line + '\n'
    
    # Gather all other relevant content lines
    for i, line in enumerate(lines):
        # Skip title, comments, horizontal rules, and the mermaid block itself
        if line.startswith('# '): continue
        if line.strip().startswith('<!--'): continue
        if line.strip() == '---': continue
        if mermaid_block_start_line <= i <= mermaid_block_end_line: continue
        
        # Skip "組織" and "職位" lines
        if "組織:" in line or "職位:" in line:
            continue

        # Keep the line if it has content
        if line.strip():
            other_content_lines.append(line)

    # --- 2. Create the Presentation ---
    
    prs = Presentation()
    # Set slide aspect ratio to 16:10 (e.g., 12 inches x 7.5 inches)
    prs.slide_width = Inches(12)
    prs.slide_height = Inches(7.5)

    # Use "Title and Content" layout
    slide_layout = prs.slide_layouts[1] 
    slide = prs.slides.add_slide(slide_layout)
    
    # Set the slide title
    slide.shapes.title.text = title
    # Apply font to title
    slide.shapes.title.text_frame.paragraphs[0].font.name = "Microsoft YaHei"
    
    # Get the body text frame
    body_shape = slide.placeholders[1]
    tf = body_shape.text_frame
    tf.clear()  # Clear default text
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    # --- Handle Mermaid Gantt Chart ---
    gantt_image_path = None
    image_height_in_inches = Inches(3.5) # Adjusted height for 16:10
    image_top_in_inches = Inches(1.2) # Adjusted top position

    if mermaid_code.strip():
        print("Attempting to render Mermaid Gantt chart to image...")
        temp_svg_path = None
        temp_png_path = None
        try:
            # Create temporary file paths
            with tempfile.NamedTemporaryFile(delete=False, suffix=".svg") as temp_svg_file:
                temp_svg_path = temp_svg_file.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_png_file:
                temp_png_path = temp_png_file.name

            # Inject font family into Mermaid code
            # This ensures Mermaid.js uses a Chinese-compatible font during SVG generation
            mermaid_code_with_font = f"%%{{init: {{'theme': 'default', 'fontFamily': 'Microsoft YaHei, sans-serif'}}}}%%\n{mermaid_code}"

            # 1. Generate SVG from Mermaid code directly to file
            mermaid.Mermaid(mermaid_code_with_font).to_svg(path=temp_svg_path)
            
            # 2. Convert SVG to PNG
            cairosvg.svg2png(url=temp_svg_path, write_to=temp_png_path)
            gantt_image_path = temp_png_path
            print(f"Mermaid Gantt chart rendered to: {gantt_image_path}")

        except Exception as e:
            print(f"Error rendering Mermaid chart to image: {e}")
            print("Falling back to text version of Gantt chart.")
            # Fallback to text if image generation fails
            p_mermaid = tf.add_paragraph()
            run = p_mermaid.add_run()
            run.text = mermaid_code
            run.font.size = Pt(12)
            run.font.name = 'Microsoft YaHei' # Use a Chinese font for the chart text
        finally:
            if temp_svg_path and os.path.exists(temp_svg_path):
                os.remove(temp_svg_path)
            # temp_png_path will be cleaned up after prs.save()

    if gantt_image_path:
        # Position the image at the top of the slide
        left = Inches(0.5)
        top = image_top_in_inches
        width = Inches(11) # Adjusted width for 16:10
        height = image_height_in_inches

        pic = slide.shapes.add_picture(gantt_image_path, left, top, width=width, height=height)
        
        # Adjust the text frame position to be below the image
        tf.top = top + height + Inches(0.2) # 0.2 inch padding
        tf.height = prs.slide_height - tf.top - Inches(0.5) # Remaining height with bottom margin

    # Add a separator line if an image was inserted, or if falling back to text
    if gantt_image_path or mermaid_code.strip():
        p_sep = tf.add_paragraph()
        run = p_sep.add_run()
        run.text = '\n' + ('-' * 40) + '\n'
        run.font.size = Pt(12)
        run.font.name = 'Microsoft YaHei' # Apply font to separator

    # --- Add today's date as the first line of the first content level ---
    today_date_str = date.today().strftime('%Y-%m-%d')
    p_date = tf.add_paragraph()
    run_date = p_date.add_run()
    run_date.text = f"Date: {today_date_str}"
    run_date.font.size = Pt(12)
    run_date.font.name = 'Microsoft YaHei'
    p_date.level = 0 # First content level

    # Add all other content
    for line_text in other_content_lines:
        p = tf.add_paragraph()
        
        cleaned_text = line_text.strip() # Remove all leading/trailing whitespace

        # Remove Links from Sub-items: finds [text](url) and replaces it with just text
        cleaned_text = re.sub(r'\[(.*?)\]\(.*?\)', r'\1', cleaned_text)

        run = p.add_run() # Add a run to the paragraph
        run.font.size = Pt(12)
        run.font.name = 'Microsoft YaHei'
        
        # Bold and Level 1 for Status Headings
        if cleaned_text.endswith('Projects'): # Heuristic to identify status headings
            run.text = cleaned_text
            run.font.bold = True
            p.level = 0 # Highest level
        elif cleaned_text.startswith('* '):
            run.text = cleaned_text.lstrip('* ').strip() # Remove bullet and any extra spaces
            p.level = 2
        elif cleaned_text.startswith('- '):
            run.text = cleaned_text.lstrip('- ').strip() # Remove bullet and any extra spaces
            p.level = 1
        else:
            run.text = cleaned_text
            p.level = 0
        
    # --- 3. Add Slide 2 (VS Code Dark Preview Style) ---
    base_output_name = os.path.splitext(output_path)[0]
    dark_png_path = base_output_name + '_dark.png'
    default_png_path = base_output_name + '.png'
    print("Generating Slide 2 preview image (VS Code Dark style)...")
    if render_preview_image(md_content, dark_png_path, theme="dark"):
        # Keep default .png as dark version as well
        shutil.copyfile(dark_png_path, default_png_path)
        slide2_layout = prs.slide_layouts[6] # Blank slide layout
        slide2 = prs.slides.add_slide(slide2_layout)
        try:
            background = slide2.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(30, 30, 30)
        except Exception:
            pass
        slide2.shapes.add_picture(dark_png_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
        print(f"Slide 2 added (VS Code Dark).")
        print(f"Saved: {dark_png_path}")

    # --- 4. Add Slide 3 (GitHub Light Preview Style) ---
    light_png_path = base_output_name + '_light.png'
    print("Generating Slide 3 preview image (GitHub Light style)...")
    if render_preview_image(md_content, light_png_path, theme="light"):
        slide3_layout = prs.slide_layouts[6] # Blank slide layout
        slide3 = prs.slides.add_slide(slide3_layout)
        try:
            background = slide3.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(255, 255, 255)
        except Exception:
            pass
        slide3.shapes.add_picture(light_png_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
        print(f"Slide 3 added (GitHub Light).")
        print(f"Saved: {light_png_path}")

    print(f"Saving presentation to: {output_path}")
    prs.save(output_path)

    # Clean up the generated PNG file after saving the presentation
    if gantt_image_path and os.path.exists(gantt_image_path):
        os.remove(gantt_image_path)


def main():
    """
    Main function to find the markdown file, parse it, and generate a presentation.
    """
    base_path = os.path.dirname(os.path.abspath(__file__))
    people_path = os.path.join(base_path, 'people')

    try:
        md_file = glob.glob(os.path.join(people_path, '*.md'))[0]
    except IndexError:
        print("Error: No markdown file found in the /people directory.")
        return

    print(f"Processing file: {md_file}")

    with open(md_file, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    output_filename = os.path.splitext(os.path.basename(md_file))[0] + '.pptx'
    output_path = os.path.join(people_path, output_filename)
    
    create_single_slide_presentation(md_content, output_path)
    print("Presentation created successfully.")

if __name__ == '__main__':
    main()