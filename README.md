# Financial Report Automation


This project is a financial report automation tool. It is designed to scrape financial data from the web and generate reports in PDF format. After the reports are generated, they're sent sent to a list of recipients via email. 


It was created as part of a freelance project for a client who wanted to automate the process of generating a financial reports for his company. The report was generated dayly and it was a very time consuming process. Also there were different formats for the report, depending on how close to the expiration date of the contracts were. The client wanted to automate the process so that he could save time and focus on other important tasks. For privacy reasons, the code is not exactly the same as the one used for the client, but it is very similar. Also the reports were in German, so the translation may not be perfect.


To scrape the data I used the following libraries:
- [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
- [Requests](https://docs.python-requests.org/en/master/)

To generate the PDF I used html and css for the layout and [pyhtml2pdf](https://pypi.org/project/pyhtml2pdf/) to convert the html to pdf, since it supports svg images, unlike other more popular libraries at the time.


To process the data I used the following libraries:
- [Pandas](https://pandas.pydata.org/)
- [Numpy](https://numpy.org/)
- [Scipy](https://www.scipy.org/)

To generate the graphs and tables images I used  [Plotly](https://plotly.com/python/) since it is very customizable and also supports svg images.

To send the emails I used the following libraries when not using the Outlook app:

- [smtplib](https://docs.python.org/3/library/smtplib.html) 
- [email](https://docs.python.org/3/library/email.examples.html).

When using the Outlook (to avoid the authentication process) I used the following libraries:

- [win32com.client](https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem)


<h1 align="center">
    <img src="https://raw.githubusercontent.com/Vinicius667/FinancialReportAutomation/main/src/images/report_example.png" 
    width="400"
    height="500" 
    />
</h1>