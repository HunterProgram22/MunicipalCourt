import urllib
import docx
import datetime
import requests
import json

from flatten_json import flatten
from loguru import logger

import docx2pdf
from docx.oxml.shared import OxmlElement, qn
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
    rPr_element = docx.oxml.shared.OxmlElement("w:rPr")

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr_element)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


VIDEO_PATH = "V:\\COURTROOM VIDEO PROCEEDINGS - JULY 2021 TO PRESENT\\"
BEARER_TOKEN = "YjI3OGFiMWItNGIzZS00NTIxLWI0NzktZDI1YzJhZDUxYjM2ODA0ZDY3NzUtOGEy_PF84_6b8180ed-7e73-4208-90f5-b67a07de84ac"


@logger.catch
def main():
    month = input("Enter month of video proceedings (i.e. 'September'):")
    year = input(str("Enter year of video proceedings (i.e. '2021'):"))
    start_date = input(
        str("Enter the date (YYYY-MM-DD) of the first day of recordings to download:")
    )
    end_date = input(
        str(
            "Enter the date (YYYY-MM-DD) of the last day of recordings to download"
            + " - download does not include the last day:"
        )
    )
    URL_string_delaware = (
        "https://webexapis.com/v1/recordings?max=100&from="
        + "{from_date}&to={to_date}&siteUrl={site_url}&hostEmail={email}".format(
            max_records="100",  # There is something up with max_records in the format
            from_date=start_date,
            to_date=end_date,
            site_url="delawareohio.webex.com",
            email="jkudela@municipalcourt.org",
        )
    )

    URL_string_municipal = (
        "https://webexapis.com/v1/recordings?max=100&from="
        + "{from_date}&to={to_date}&siteUrl={site_url}&hostEmail={email}".format(
            max_records="100",  # There is something up with max_records in the format
            from_date=start_date,
            to_date=end_date,
            site_url="municipalcourt.webex.com",
            email="jkudela@municipalcourt.org",
        )
    )
    # logger.info(URL_string)

    location = "Delaware City Webex"
    headers = {
        "Authorization": "Bearer {bearer_token}".format(bearer_token=BEARER_TOKEN)
    }
    response_delaware = requests.get(
        url=URL_string_delaware,
        headers=headers,
        stream=True,
    )

    response_municipal = requests.get(
        url=URL_string_municipal,
        headers=headers,
        stream=True,
    )
    # logger.info(response)
    response_data_delaware = response_delaware.json()
    response_data_municipal = response_municipal.json()
    # logger.info(response_data)
    response_items_delaware = response_data_delaware["items"]
    response_items_municipal = response_data_municipal["items"]

    mydoc = docx.Document()
    start_date_list = start_date.split("-")
    heading_date = (
        start_date_list[1] + "-" + start_date_list[2] + "-" + start_date_list[0]
    )
    heading = mydoc.add_heading("Court Video Proceedings " + heading_date + "\n")
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    mydoc_paragraph = mydoc.add_paragraph()
    for video_info in response_items_delaware:
        playback_url = video_info["playbackUrl"]
        title = video_info["topic"]
        print(title)
        print(playback_url)
        run = mydoc_paragraph.add_run(title + "\n")
        run.bold
        run.bold = True
        run_link = add_hyperlink(mydoc_paragraph, playback_url, "Video Link")
        mydoc_paragraph.add_run("\n\n")
    for video_info in response_items_municipal:
        playback_url = video_info["playbackUrl"]
        title = video_info["topic"]
        print(title)
        print(playback_url)
        run = mydoc_paragraph.add_run(title + "\n")
        run.bold
        run.bold = True
        run_link = add_hyperlink(mydoc_paragraph, playback_url, "Video Link")
        mydoc_paragraph.add_run("\n\n")

    document_name = "Courtroom_Proceedings_" + heading_date + ".docx"
    mydoc.save(VIDEO_PATH + month + " " + year + "\\" + document_name)
    docx2pdf.convert(VIDEO_PATH + month + " " + year + "\\" + document_name)


main()
