# strike
On-refresh tallying of color coded Google Docs bulleted lists.


### This thing has certain quirks. Here is a list of things I know about: ###

- Apps Script is not returning strikethrough values for text that is stricken via checkbox, meaning only manual/alt+shift+5 strikethroughs work for now. Using the checkboxes would mess up the tally.

- It is critical that at least two full empty lines are added at the end of the bulleted list. Adding the table every time used to add an unwanted space so each time the last character is cleared. For some reason, this seems to malfunction if the above criterion isn't met.

- Properties evaluated at 0 index: The bold, strikethrough, and background color properties are only checked at the first character of each list item. This means that the entire line formatting should match that of the first character for proper results. The reason is that newline characters complicate attribute checking such that properties that appear like they should return a value return null. e.g., background color appears to cover the whole list item by default but actually doesn't extend to the carriage return unless there is another text line beneath. This would cause multiple background color values which gets returned as null.

- Corollary to the above: Empty bullet points (bulleted lines with no text) are going to cause an error and fail to produce the tally. This is an easy fix and I am likely to fix it soon.

- The code assumes there are no tables in the document. Any tables in a document will get deleted. Could fix with relative ease if necessary.
