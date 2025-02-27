import os
import sys
import win32com.client
from win32com.client import constants
import re

def create_word_bookmarks(docx_path, target_words):
    """
    Creates bookmarks at all occurrences of the specified words in a Word document
    and returns a list of links to those bookmarks.
    
    Args:
        docx_path (str): Path to the Word document
        target_words (list): List of words to find and bookmark
    
    Returns:
        str: Markdown formatted text with links to all bookmarks
    """
    # Convert to absolute path
    docx_path = os.path.abspath(docx_path)
    
    if not os.path.exists(docx_path):
        return f"Error: File not found at {docx_path}"
    
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
            word_app.Selection.HomeKey(Unit=constants.wdStory)
            
            # Counter for occurrences
            count = 0
            
            # Loop through all occurrences
            while word_app.Selection.Find.Execute():
                count += 1
                # Get current selection information
                selected_text = word_app.Selection.Text
                
                # Create unique bookmark name
                bookmark_name = f"{word.replace(' ', '_')}_{count}"
                
                # Check if bookmark already exists
                bookmark_exists = False
                for i in range(1, doc.Bookmarks.Count + 1):
                    if doc.Bookmarks.Item(i).Name == bookmark_name:
                        bookmark_exists = True
                        break
                
                if not bookmark_exists:
                    # Add bookmark
                    doc.Bookmarks.Add(bookmark_name)
                
                # Get paragraph number and position for reporting
                para_index = word_app.Selection.Information(constants.wdActiveEndAdjustedPageNumber)
                position = word_app.Selection.Information(constants.wdFirstCharacterColumnNumber)
                
                # Store information
                results[word].append({
                    "bookmark_name": bookmark_name,
                    "text": selected_text.strip(),
                    "paragraph": para_index,
                    "position": position
                })
                
                # Move selection to end of found text to continue search
                word_app.Selection.MoveRight(Unit=constants.wdCharacter, Count=1)
        
        # Save document with bookmarks
        output_path = docx_path.replace(".docx", "_bookmarked.docx")
        doc.SaveAs(output_path)
        doc.Close()
        
        # Generate markdown report
        report = ["# Word Bookmark Links\n"]
        report.append(f"Document: {os.path.basename(docx_path)}\n")
        report.append("The following links will work when viewing the saved document: "
                      f"{os.path.basename(output_path)}\n")
        
        for word in target_words:
            report.append(f"## Links for '{word}':\n")
            if results[word]:
                for i, bookmark in enumerate(results[word], 1):
                    bookmark_link = f"#{bookmark['bookmark_name']}"
                    report.append(f"{i}. [{bookmark['text']}]({bookmark_link}) "
                                 f"(Page {bookmark['paragraph']}, Position {bookmark['position']})")
            else:
                report.append("No occurrences found.")
            report.append("")
        
        print(f"Document saved with bookmarks as: {output_path}")
        return "\n".join(report)
    
    except Exception as e:
        return f"Error: {str(e)}"
    
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
        python word_bookmarker.py "path/to/document.docx" "word1,word2,word3"
    """
    if len(sys.argv) < 3:
        print("Usage: python word_bookmarker.py \"path/to/document.docx\" \"word1,word2,word3\"")
        return
    
    docx_path = sys.argv[1]
    target_words = [word.strip() for word in sys.argv[2].split(",")]
    
    result = create_word_bookmarks(docx_path, target_words)
    
    # Save report to file
    report_path = os.path.join(os.path.dirname(docx_path), "bookmark_links.md")
    with open(report_path, "w") as f:
        f.write(result)
    
    print(f"Report saved to: {report_path}")
    print("\n" + result)

if __name__ == "__main__":
    main()
