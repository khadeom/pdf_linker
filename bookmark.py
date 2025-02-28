import os
import sys
import win32com.client
import re
import webbrowser
from urllib.parse import quote

def create_word_bookmarks_with_links(docx_path, target_words):
    """
    Creates bookmarks in a Word document and generates a text file with clickable links
    that will open the document and navigate to the bookmarked words.
    
    Args:
        docx_path (str): Path to the Word document
        target_words (list): List of words to find and bookmark
    
    Returns:
        str: Path to the generated links file
    """
    # Convert to absolute path
    docx_path = os.path.abspath(docx_path)
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found at {docx_path}")
        return None
    
    # Start Word application
    print("Starting Word application...")
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False  # Run in background
    
    try:
        # Open the document
        print(f"Opening document: {docx_path}")
        doc = word_app.Documents.Open(docx_path)
        
        # Dictionary to track results
        results = {word: [] for word in target_words}
        
        # Process each target word
        for word in target_words:
            print(f"Processing word: '{word}'")
            
            # Setup the Find object
            find_obj = word_app.Selection.Find
            find_obj.ClearFormatting()
            find_obj.Text = word
            find_obj.MatchWholeWord = True
            find_obj.MatchCase = False
            
            # Reset cursor to beginning of document
            word_app.Selection.Start = 0
            
            # Counter for occurrences
            count = 0
            
            # Loop through all occurrences
            while word_app.Selection.Find.Execute():
                count += 1
                # Get current selection information
                selected_text = word_app.Selection.Text
                
                # Create unique bookmark name
                bookmark_name = f"{word.replace(' ', '_')}_{count}"
                
                # Check if bookmark already exists and delete it if it does
                try:
                    existing = doc.Bookmarks(bookmark_name)
                    existing.Delete()
                except:
                    pass
                
                # Add bookmark
                doc.Bookmarks.Add(bookmark_name)
                
                # Get information for reporting
                current_pos = word_app.Selection.Start
                para_num = word_app.Selection.Information(3)  # 3 = paragraph number
                
                # Store information
                results[word].append({
                    "bookmark_name": bookmark_name,
                    "text": selected_text.strip(),
                    "paragraph": para_num,
                    "position": current_pos
                })
                
                # Move selection to end of found text to continue search
                word_app.Selection.Start = word_app.Selection.End
        
        # Save document with bookmarks
        output_path = docx_path.replace(".docx", "_bookmarked.docx")
        doc.SaveAs(output_path)
        doc.Close()
        
        # Generate hyperlinks file
        links_path = os.path.join(os.path.dirname(docx_path), "word_links.txt")
        with open(links_path, "w") as f:
            f.write(f"LINKS TO BOOKMARKED WORDS IN {os.path.basename(output_path)}\n")
            f.write("-" * 60 + "\n\n")
            
            for word in target_words:
                f.write(f"Links for '{word}':\n")
                if results[word]:
                    for i, bookmark in enumerate(results[word], 1):
                        # Create a file:// URL with the bookmark
                        file_url = f"file:///{output_path.replace('\\', '/')}#{bookmark['bookmark_name']}"
                        
                        # Write the link
                        f.write(f"{i}. {bookmark['text']} (Paragraph {bookmark['paragraph']})\n")
                        f.write(f"   {file_url}\n\n")
                else:
                    f.write("   No occurrences found.\n\n")
        
        print(f"Document saved with bookmarks as: {output_path}")
        print(f"Links saved to: {links_path}")
        return links_path
    
    except Exception as e:
        print(f"Error: {str(e)}")
        return None
    
    finally:
        # Close Word application
        try:
            word_app.Quit()
        except:
            pass

def main():
    """
    Main function to parse command line arguments and run the bookmark creator.
    
    Usage: 
        python word_links.py "path/to/document.docx" "word1,word2,word3"
    """
    if len(sys.argv) < 3:
        print("Usage: python word_links.py \"path/to/document.docx\" \"word1,word2,word3\"")
        return
    
    docx_path = sys.argv[1]
    target_words = [word.strip() for word in sys.argv[2].split(",")]
    
    links_file = create_word_bookmarks_with_links(docx_path, target_words)
    
    if links_file:
        # Open the links file automatically
        print(f"Opening links file: {links_file}")
        os.startfile(links_file)

if __name__ == "__main__":
    main()