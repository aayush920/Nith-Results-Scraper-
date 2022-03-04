# Nith-Results-Scraper
Python script to scrape the [NITH Result](https://nithp.herokuapp.com/result/) website and retrieve the result of a particular student.<br>
Takes a Roll. No. and an email address as input ,generates the appropriate url and retrieves all the data and writes it into a newly created excel file using [xlsxwriter](https://xlsxwriter.readthedocs.io/worksheet.html) module and sends it as an attachment to the email using [smtplib](https://docs.python.org/3/library/smtplib.html) module.<br><br><br>

The sender email address, email pasword, file location are stored as environment variables in the local system and retrieved using the [OS](https://docs.python.org/3/library/os.html) module so that they are not visible in the script.<br>
<img width="1000" alt="temp1" src="https://user-images.githubusercontent.com/76609501/156235700-e5f912a2-8836-48d6-9af4-041fe6f4a491.png"><br><br>

<b>Sample Result</b>:<br>
<img width="1000" alt="image-1" src="https://user-images.githubusercontent.com/76609501/156234954-0911de57-061e-4c8a-8bc9-318c1f28a138.png">


