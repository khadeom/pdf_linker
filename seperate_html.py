from bs4 import BeautifulSoup
import os

# Load the HTML file
html_file_path = "index.html"  # Change this to your file
with open(html_file_path, "r", encoding="utf-8") as file:
    soup = BeautifulSoup(file, "html.parser")

# Create an output directory
output_dir = "split_html_pages"
os.makedirs(output_dir, exist_ok=True)

# Extract CSS styles
styles = soup.find_all("style")
css_content = "\n".join([str(style) for style in styles])

# Extract inline scripts (excluding external scripts)
scripts = soup.find_all("script")
js_content = "\n".join([str(script) for script in scripts if not script.has_attr("src")])

# Identify all major content sections dynamically
valid_tags = ["section", "div", "article", "main"]  # Add more if needed
containers = [tag for tag in soup.find_all(valid_tags) if tag.parent.name == "body"]  # Ensure top-level elements

if not containers:
    print("No valid sections found. Saving full body as a single page.")
    containers = [soup.body]  # If no specific sections found, use entire body

# Save each extracted part as a separate HTML file
for idx, container in enumerate(containers):
    # Wrap extracted content in a full HTML document
    page_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Page {idx+1}</title>
        {css_content}  <!-- Inject CSS -->
    </head>
    <body>
        {container}  <!-- Inject extracted content -->
        {js_content}  <!-- Inject inline JS -->
    </body>
    </html>
    """

    # Save each section as an HTML file
    file_path = os.path.join(output_dir, f"page_{idx+1}.html")
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(page_content)

    print(f"Saved: {file_path}")

print("Splitting completed!")