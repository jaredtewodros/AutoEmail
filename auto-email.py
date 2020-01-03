#import win32 for email, matplotlib for radar chart and pandas
import win32com.client as win32
import pandas as pd
import matplotlib.image as image
import matplotlib.pyplot as plt
import matplotlib.style as style
import numpy as np
import matplotlib.patches as mpatches
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
import os

#read data from csv, skip the first two rows as those names are not conventional
data = pd.read_csv('BookPythonTest1.csv',skiprows=2)
data = pd.DataFrame(data)



# Sending PDF using reportlab
def sendReportLab():
    # Create PDF
    #styles = getSampleStyleSheet()
    doc = SimpleDocTemplate("Results.pdf", pagesize=letter,
                            rightMargin=72, leftMargin=72,
                            topMargin=12, bottomMargin=18)

    # c = canvas.Canvas("Results.pdf")
   # c.setFillColorRGB(1, 0, 0)
   # c.rect(5, 5, 652, 792, fill=1)
    Story = []
    logo = "xScion Logo Alt.png"
    magName = "Pythonista"

    #canvas.rect(0,0,8*inch,11*inch,stroke=0,fill=1)

    im = Image(logo,3*inch,1*inch)
    Story.append(im)

    # Add Styles

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY))
    styles.add(ParagraphStyle(name='Content',
                              fontFamily='Helvetica',
                              fontSize=8,
                              TextColor=colors.HexColor("#E57200")))

    # Create Some Space between the Logo and the remaining text
    Story.append(Spacer(1, 24))
    x = str(row['First Name*'])
    ptext = 'Hello ' + x + ','
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 12))

    x = str(row["Organization*"])
    ptext = 'Thank you for taking the Agile Data Maturity Model flash survey. Below is a high-level view of where '+x+" stands in its Data Maturity."
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 12))
    ptext = 'This image shows how advanced your organization is on 13 different data capabilities, ranked each from Foundational, Intermediate and Advanced.'
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 12))
    RadarImage = Image("radarChart.png", 8 * inch, 4.5 * inch)
    Story.append(RadarImage)
    Story.append(Spacer(1, 12))
    if (row["Score"] > 30):
        temp = "Advanced"
    elif (row["Score"] > 15):
        temp = "Intermediate"
    else:
        temp = "Foundational"
    ptext = "Overall, you scored a " + str(row["Score"]) + " out of 45, which ranks as an overall " + temp +" level."
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 12))


    ptext = "Based on your Data Assessment, our Principal Strategist has determined the following:"
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 12))

    ourWords = [
        row["Question 1 Analysis"],
        row["Question 2 Analysis"],
        row["Question 3 Analysis"],
        row["Question 4 Analysis"],
        row["Question 5 Analysis"],
        row["Question 6 Analysis"],
        row["Question 7 Analysis"],
        row["Question 8 Analysis"],
        row["Question 9 Analysis"],
        row["Question 10 Analysis"],
        row["Question 11 Analysis"],
        row["Question 12 Analysis"],
        row["Question 13 Analysis"],
        row["Question 14 Analysis"],
        row["Question 15 Analysis"],
    ]
    x = str(row["Organization*"])
    ourQuestions = [
        "Data is key to " + x + " Corporate Strategic Plan?",
        x + ' could be considered a "data-driven" company because it exhibits all three below traits: 1. Sees data as its core competency, including leverages and monetizes data to generate value for stakeholders. 2. Ensures every person who can use data to make better decisions, has access to the data they need when they need it. 3. Relentlessly measures and monitors the pulse of the business and are doing so in a continuous, often-automated manner.',
        "Using analytics to transform " + x + " is an executive management goal.",
        x + " has a culture of accountability and shared responsibility for managing corporate data assets.",
        x + " has a corporate Business Glossary and standard approach for managing metadata.",
        x + " has developed policies and standards for ensuring that the right people have access to the right data at the right time.",
        x + " has developed Conceptual, Logical, and Physical models for its Data Architecture.",
        x + " has a dedicated Data Architecture Group.",
        x + " considers changes to its Data Model as part of the System Development",
        x + " has a Data Quality Engineering Group.",
        x + " has developed Quality Assurance processes, which are actively integrated and monitored across all business portfolios, to measure data quality and timeliness.",
        x + " has a formal, documented Disaster Recovery plan for its data.",
        x + " has a Big Data Operations Group.",
        x + " processes structured and unstructured data.",
        x + " has invested in data technologies/tools, such as Predictive Analytics, Machine Learning, and Artificial Intelligence."
    ]
    theAnswers = [
        row["Q8"],
        row["Q9"],
        row["Q10"],
        row["Q11"],
        row["Q12"],
        row["Q13"],
        row["Q14"],
        row["Q15"],
        row["Q16"],
        row["Q17"],
        row["Q18"],
        row["Q19"],
        row["Q20"],
        row["Q21"],
        row["Q22"],
    ]

    for x,y,z in zip(ourWords, ourQuestions,theAnswers):
        a = x
        x = str(x)
        if x != "nan":
            #global emailText
            #emailText = emailText,"<li>", y,"<br>", z,"<br/>",x
            #print(y,a,z)
            ptext = y
            Story.append(Paragraph(ptext, styles["Normal"],bulletText='-'))
            Story.append(Spacer(1, 3))
            ptext = "Your Answer: "+ z
            Story.append(Paragraph(ptext, styles["Normal"]))
            Story.append(Spacer(1, 3))
            ptext = x
            Story.append(Paragraph(ptext, styles["Normal"]))
            Story.append(Spacer(1, 12))
    org = str(row["Organization*"] )
    ptext = "This flash survey is just a small glimpse into your organization’s data people, process and technology. Our full Agile Data Maturity Model includes an in-depth analysis across your organization, a roadmap for improved maturity and support to help you get there. We’d love to help " + \
            org+ " on its path forward. Please feel free to contact us anytime at " + "<a href='mailto:admm@xscion.com'>admm@xscion.com</a>" + " or (571)-425-4800."
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 12))

    ptext = 'Thanks,'
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 3))
    ptext = 'xScion Solutions Team'
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 3))
    ptext = "<a href='mailto:admm@xscion.com'>admm@xscion.com</a>"
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 3))
    ptext = "Learn more: <a href='https://www.xscion.com'>www.xscion.com</a>"
    Story.append(Paragraph(ptext, styles["Normal"]))
    Story.append(Spacer(1, 3))
    doc.build(Story)


#iterrate through all rows, thus sending emails to all users
for index, row in data.iterrows():
    style.use('dark_background')
    im = image.imread('xScion Logo.png')

    #Set up Labels and Data
    labels = np.array(["Program Goals", "Program Scope", "Business Case", "Ownership/Stewardship", "Metadata Management",
    "Quality & Security", "Data Architecture", "Data Integration", "Quality Engineering", "Quality Assurance",
    "Disaster Recovery", "Data Science", "Predictive Analytics"])

    otherStats = [row["Program Goals"], row["Program Scope"], row["Business Case & Program Funding"],
                  row["Ownership & Stewardship"],row["Business Glossary & Metadata Management"],
    row["Quality & Security"],row["Data Architecture"],row["Data Integration"],row["Quality Engineering Staff/Skills"],
                  row["Process Quality Assurance"],
    row["Disaster Recovery"], row["Data Science"],row["Predictive Analytics, Machine Learning & Artificial Intelligence"]]

    #Calculations for graph
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False)
    angles = np.concatenate((angles, [angles[0]]))
    fig=plt.figure(figsize=(16,9))
    otherStats = np.concatenate((otherStats, [otherStats[0]]))

    #Set up Font Sizes
    plt.rcParams.update({'font.size': 16})
    #Show plot
    ax2 = fig.add_subplot(1, 1, 1, polar=True)
    ax2.plot(angles, otherStats, 'o-', linewidth=2, color="blue", label=row["Organization*"])
    ax2.fill(angles, otherStats, alpha=0.2, color="blue")
    ax2.set_thetagrids(angles * 180 / np.pi, labels)
    ax2.figure.figimage(im, 40, 40, alpha=.9, zorder=1)
    ax2.axes.get_yaxis().set_ticklabels(["","Fdn.","","Inter.","","Adv."])
    plt.gca().get_xticklabels()[0].set_color('#bfbfbf')
    plt.gca().get_xticklabels()[1].set_color('#bfbfbf')
    plt.gca().get_xticklabels()[2].set_color('#bfbfbf')
    plt.gca().get_xticklabels()[3].set_color('#CC6600')
    plt.gca().get_xticklabels()[4].set_color('#CC6600')
    plt.gca().get_xticklabels()[5].set_color('#CC6600')
    plt.gca().get_xticklabels()[6].set_color('#00ff00')
    plt.gca().get_xticklabels()[7].set_color('#00ff00')
    plt.gca().get_xticklabels()[8].set_color('#fe1b07')
    plt.gca().get_xticklabels()[9].set_color('#fe1b07')
    plt.gca().get_xticklabels()[10].set_color('#fe1b07')
    plt.gca().get_xticklabels()[11].set_color('#ff00e1')
    plt.gca().get_xticklabels()[12].set_color('#ff00e1')
    s_patch = mpatches.Patch(color='#bfbfbf', label='Strategy')
    g_patch = mpatches.Patch(color='#CC6600', label='Governance')
    a_patch = mpatches.Patch(color='#00ff00', label='Architecture')
    qe_patch = mpatches.Patch(color='#fe1b07', label='Quality Engr.')
    an_patch = mpatches.Patch(color='#ff00e1', label='Analytics')
    plt.legend(bbox_to_anchor=(1.5, 1.1), handles=[s_patch,g_patch,a_patch,qe_patch,an_patch])
    ax2.set_rlabel_position(270)
    ax2.set_title(row["Organization*"])
    #Save image and wait
    plt.savefig('radarChart.png')
    import time
    #time.sleep(10)

    #Email Contents
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    ##Change mail
    mail.To = "jaredtewodros@gmail.com"
    mail.Subject = 'Your Results: ADMM Flash Survey'


    ourWords = [
        row["Question 1 Analysis"],
        row["Question 2 Analysis"],
        row["Question 3 Analysis"],
        row["Question 4 Analysis"],
        row["Question 5 Analysis"],
        row["Question 6 Analysis"],
        row["Question 7 Analysis"],
        row["Question 8 Analysis"],
        row["Question 9 Analysis"],
        row["Question 10 Analysis"],
        row["Question 11 Analysis"],
        row["Question 12 Analysis"],
        row["Question 13 Analysis"],
        row["Question 14 Analysis"],
        row["Question 15 Analysis"],
    ]
    x = str(row["Organization*"])
    ourQuestions = [
        'Data is key to ' + x + ' Corporate Strategic Plan?',
        x + ' could be considered a "data-driven" company because it exhibits all three below traits: 1. Sees data as its core competency, including leverages and monetizes data to generate value for stakeholders. 2. Ensures every person who can use data to make better decisions, has access to the data they need when they need it. 3. Relentlessly measures and monitors the pulse of the business and are doing so in a continuous, often-automated manner.',
        "Using analytics to transform " + x + " is an executive management goal.",
        x + " has a culture of accountability and shared responsibility for managing corporate data assets.",
        x + " has a corporate Business Glossary and standard approach for managing metadata.",
        x + " has developed policies and standards for ensuring that the right people have access to the right data at the right time.",
        x + " has developed Conceptual, Logical, and Physical models for its Data Architecture.",
        x + " has a dedicated Data Architecture Group.",
        x + " considers changes to its Data Model as part of the System Development",
        x + " has a Data Quality Engineering Group.",
        x + " has developed Quality Assurance processes, which are actively integrated and monitored across all business portfolios, to measure data quality and timeliness.",
        x + " has a formal, documented Disaster Recovery plan for its data.",
        x + " has a Big Data Operations Group.",
        x + " processes structured and unstructured data.",
        x + " has invested in data technologies/tools, such as Predictive Analytics, Machine Learning, and Artificial Intelligence."
    ]
    theAnswers = [
        row["Q8"],
        row["Q9"],
        row["Q10"],
        row["Q11"],
        row["Q12"],
        row["Q13"],
        row["Q14"],
        row["Q15"],
        row["Q16"],
        row["Q17"],
        row["Q18"],
        row["Q19"],
        row["Q20"],
        row["Q21"],
        row["Q22"],
    ]

    emailText = '<ul>'
    for x,y,z in zip(ourWords, ourQuestions,theAnswers):
        a = x
        x = str(x)
        if x != "nan":
            x = str(x)
            y = str(y)
            z = str(z)
            emailText = str(emailText)
            emailText = emailText + '<li>' + y + '<br> Your Answer: '+ z+'<br>'+x+'<br></li>'
    #emailText += '</ul>'
    print(emailText)
    #emailText = str(emailText)
    firstName = str(row['First Name*'])
    org = str(row["Organization*"])
    body = 'Hello ' + firstName +',<br><br>' + 'Thank you for taking the Agile Data Maturity Model flash survey. Attached is a high-level view of where ' + org + ' stands in its Data Maturity.<br><br>' + emailText\
           +'<br>This flash survey is just a small glimpse into your organization’s data people, process and technology. Our full Agile Data Maturity Model includes an in-depth analysis across your organization, a roadmap for improved maturity, and support to help you get there. We’d love to help '\
    +org + ' on its path forward. Also, for your viewing convenience, this email is attached as a PDF. Please feel free to contact us anytime at ' + "<a href='mailto:admm@xscion.com'>admm@xscion.com</a>  or (571)-425-4800.<br><br>" +"Thanks,<br>xScion Solutions Team<br> e: <a href='mailto:admm@xscion.com'>admm@xscion.com</a><br> Learn more: <a href='http://www.xscion.com'>www.xscion.com</a>"
    mail.HTMLBody = body

    # Create PDF
    sendReportLab()
    time.sleep(10)

    # To attach the file to the email
    # This may need to get modified depending on where this repo is cloned in
    attachment1 = "C:\\Users\Owner\\GitLab\\devops-practice\\python\\Emailing\\radarChart.png"
    attachment2 = "C:\\Users\Owner\\GitLab\\devops-practice\\python\\Emailing\\Results.pdf"
    mail.Attachments.Add(attachment1)
    mail.Attachments.Add(attachment2)
    mail.Send()

    #remove image to prevent resending the image
    os.remove('radarChart.png')
    os.remove("Results.pdf")
    time.sleep(10)
