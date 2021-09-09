import requests
import json
from pandas import json_normalize
from flatten_json import flatten, flatten_preserve_lists
import pandas as pd
import urllib
import docx
import datetime
import docx2pdf
from docx.oxml.shared import OxmlElement, qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.enum.dml import MSO_THEME_COLOR_INDEX
from docx.enum.text import WD_ALIGN_PARAGRAPH


def add_hyperlink(paragraph, url, text):
    """
    A function that places a hyperlink within a paragraph object.

    :param paragraph: The paragraph we are adding the hyperlink to.
    :param url: A string containing the required url
    :param text: The text displayed for the url
    :return: The hyperlink object
    """
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(
        url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True
    )

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement("w:hyperlink")
    hyperlink.set(
        docx.oxml.shared.qn("r:id"),
        r_id,
    )

    # Create a w:r element
    new_run = docx.oxml.shared.OxmlElement("w:r")

    # Create a new w:rPr element
    rPr = docx.oxml.shared.OxmlElement("w:rPr")

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)

    return hyperlink


BEARER_TOKEN = "YzJkOTNjZGQtYzA5ZS00ZDUxLWI0MGMtZjEwN2U5YjA2YWNhMjBlNTBkOGItMTRi_PF84_6b8180ed-7e73-4208-90f5-b67a07de84ac"

# URL = "https://webexapis.com/v1/recordings?max=100&from=2021-08-24&to=2021-08-27&siteUrl=delawareohio.webex.com&hostEmail=jkudela@municipalcourt.org"
URL_string = "https://webexapis.com/v1/recordings?max={max_records}&from={from_date}&to={to_date}&siteUrl={site_url}&hostEmail={email}".format(
    max_records=100,
    from_date="2021-09-07",
    to_date="2021-09-08",
    site_url="delawareohio.webex.com",
    email="jkudela@municipalcourt.org",
)

URL = URL_string

location = "Delaware City Webex"
headers = {"Authorization": "Bearer {bearer_token}".format(bearer_token=BEARER_TOKEN)}
response = requests.get(
    url=URL,
    headers=headers,
    stream=True,
)
response_data = response.json()
print(response_data)
response_items = response_data["items"]
print(response_items)

mydoc = docx.Document()
heading_date = "09-07-2021"  # str(datetime.datetime.now().strftime("%m-%d-%Y"))
heading = mydoc.add_heading("Court Video Proceedings " + heading_date + "\n")
heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
paragraph = mydoc.add_paragraph()
# download_url = response_items[0]['downloadUrl']
for video_info in response_items:
    playback_url = video_info["playbackUrl"]
    title = video_info["topic"]
    print(title)
    print(playback_url)
    run = paragraph.add_run(title + "\n")
    run.bold
    run.bold = True
    run_link = add_hyperlink(paragraph, playback_url, "Video Link")
    paragraph.add_run("\n\n")

# mydoc.save("C:\\users\\jkudela\\Desktop\\test_download.docx")

mydoc.save(
    "V:\\COURTROOM VIDEO PROCEEDINGS - JULY 2021 TO PRESENT\\September 2021\\Courtroom_Proceedings_"
    + heading_date
    + ".docx"
)
docx2pdf.convert(
    "V:\\COURTROOM VIDEO PROCEEDINGS - JULY 2021 TO PRESENT\\September 2021\\Courtroom_Proceedings_"
    + heading_date
    + ".docx"
)
