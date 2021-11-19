from docx import Document
import json

faqpageProperties = [
  "title",
  "01question",
  "01answer",
  "02question",
  "02answer",
  "03question",
  "03answer",
  "04question",
  "04answer",
  "05question",
  "05answer",
  "06question",
  "06answer",
  "07question",
  "07answer",
  "08question",
  "08answer",
  "09question",
  "09answer",
  "10question",
  "10answer",
  "11question",
  "11answer",
  "12question",
  "12answer",
  "13question",
  "13answer",
  "14question",
  "14answer",
  "15question",
  "15answer"
]

needHelpProperties = [
  "needhelptitle",
  "needhelpcontent01",
  "needhelpcontent02",
  "needhelpcontent03",
  "needhelpcontent04",
  "needhelpcontent05",
  "needhelpcontent06"
]

# Loop through paragraph texts in FAQ section and add faqpage property to output_dict
def parseFaq(document, output_dict):
  paragraph_texts = list(map(lambda x: x.text, document.paragraphs))
  # Remove empty strings
  paragraph_texts = list(filter(lambda x: x, paragraph_texts))
  lastTableIndex = paragraph_texts.index("FormatNotFoundHTML (No match found)")
  # Some word docs have a newline after the last table
  needHelpStartIndex = lastTableIndex + 3 if paragraph_texts[lastTableIndex + 1] == '\n' else lastTableIndex + 2
  faqStartIndex = needHelpStartIndex + 7

  output_dict["faqpage"] = {}
  for prop in faqpageProperties:
    output_dict["faqpage"][prop] = paragraph_texts[faqStartIndex]
    faqStartIndex += 1
    # Some answers require multiple paragraphs
    if prop in {"03answer", "04answer", "05answer", "10answer", "12answer"}:
      output_dict["faqpage"][prop] += "<br />{}".format(paragraph_texts[faqStartIndex])
      faqStartIndex += 1
    elif prop == "07answer":
      output_dict["faqpage"][prop] += "<br /><br / >°{}".format(paragraph_texts[faqStartIndex])
      output_dict["faqpage"][prop] += "<br / >°{}".format(paragraph_texts[faqStartIndex + 1])
      output_dict["faqpage"][prop] += "<br / >°{}".format(paragraph_texts[faqStartIndex + 2])
      output_dict["faqpage"][prop] += "<br / >{}".format(paragraph_texts[faqStartIndex + 3])
      output_dict["faqpage"][prop] += "<br / >{}".format(paragraph_texts[faqStartIndex + 4])
      faqStartIndex += 5
    elif prop == "11answer":
      output_dict["faqpage"][prop] += "<br />{}".format(paragraph_texts[faqStartIndex])
      output_dict["faqpage"][prop] += "<i>{}</i>".format(paragraph_texts[faqStartIndex + 1])
      faqStartIndex += 2

  for prop in needHelpProperties:
    output_dict["faqpage"][prop] = paragraph_texts[needHelpStartIndex]
    needHelpStartIndex += 1
