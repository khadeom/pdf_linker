soup = BeautifulSoup(html_content, "html.parser")

def remove_empty_elements(tag):
    """Recursively removes elements that contain no meaningful text."""
    if tag.name:  # Ensure it's a tag, not just a string
        for child in tag.find_all(True):  # Find all child elements
            remove_empty_elements(child)  # Recursively process child elements

        # If a tag has no text (excluding spaces) and no meaningful children, remove it
        if not tag.get_text(strip=True) and not tag.find(True):  
            tag.decompose()

# Start the cleanup process from the body (avoid removing <html> and <head>)
remove_empty_elements(soup.body)

# Print the cleaned-up HTML
print(soup.prettify())