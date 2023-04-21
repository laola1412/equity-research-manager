from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib import colors


# replace in the future with the actual data
description_qrtea = 'Qurate Retail GroupSM, a holding company that owns and operates prominent brands such as QVC®, HSN®, Zulily®, Ballard Designs®, Frontgate®, Garnet Hill®, and Grandin Road®, is recognized as the largest player in the "video commerce" industry in the United States. Qurate is an integral part of the Liberty empire, led by John Malone, which includes Liberty SiriusXM (LSXMA), Liberty Broadband (LBRDA), Formula One (FWONA), the Atlanta Braves, and other significant holdings.Furthermore, Qurate\'s capital structure is noteworthy, as it comprises a diverse range of publicly traded securities. This includes $350 million in equity represented by QRTEA and QRTEB, $1.3 billion face value of 8% preferred shares listed on NASDAQ under the symbol QRTEP, and $6.9 billion worth of bonds, some of which are traded on Nasdaq (QVCC) and NYSE (QVCD).\nIn addition, Qurate\'s capital structure presents additional intriguing opportunities. For instance, the 2029 and 2030 exchangeable debentures can be converted into shares of T-Mobile (TMUS), Lumen (LUMN), and Charter Communications (CHTR) respectively. Furthermore, the unsecured bonds due in 2029 offer a substantial yield of 35% and are currently trading at 20% of face value, while other issues also demonstrate poor trading performance. Despite this, they appear to be a more sensible option compared to the subordinated preferred securities, which yield approximately 23% (or around 31% to maturity). This potentially presents an interesting but volatile arbitrage opportunity.Furthermore, Qurate retains a significant amount of cash on its balance sheet, with $1.275 billion remaining at the end of Q4-22, and an additional $182 million expected from real estate proceeds in Q1-23. Management has indicated the ability to sell more real estate if necessary, but currently, they express comfort with the existing liquidity levels. Additionally, ongoing discussions regarding business interruption insurance claims are likely to have a material impact. Notably, Qurate also holds a substantial stake in comScore (SCOR) preferred shares, acquired for $68 million in 2021.'

operations_example = 'Qurate\'s performance has been challenging, with bonds yielding over 30% indicating a poor market perception. Following the payment of $1.7 billion in dividends in the post-Covid period, Qurate experienced a warehouse fire in Rocky Mount, NC in 2022, which led to constrained inventory levels and adverse business decisions to clear stale inventory. However, management has outlined a turnaround plan called "Project Athens" with a target of achieving $300 million to $500 million of Free Cash Flow (FCF) by 2024. While Qurate reported essentially flat FCF in 2022, it was boosted by $693 million of sale-leasebacks and $280 million of insurance proceeds from the fire, but negatively impacted by $150 million to $200 million of working capital headwinds. Adjusting for these factors, the "actual" FCF was closer to negative $800 million. Qurate may leverage owned real estate and expect working capital to be a tailwind in 2023, but considerable work remains to achieve the Project Athens targets. However, these transactions may improve liquidity and optical cash flow generation, but create future earnings headwinds of $57 million (based on a 7.5% cap rate on $765 million aggregate proceeds above). If the proceeds can be used to retire debt at a discount, the return could be higher.It is crucial to monitor customer count as it took a direct hit in 2022 due to the stale inventory issue, and investing in Qurate requires faith in their ability to regain lost customers. However, given cord-cutting being cited as the primary reason for lost customers and with key metrics down from pre-Covid times, there may be doubts about a full recovery in customer count. Qurate has a strong moat, having fended off attempted entry by Amazon (AMZN) in under a year. However, current performance across categories has been challenging, providing little hope that the retention issue is limited to a subset of customers.'


def create_report(save_name, name, close, date, target_value, rating, description=description_qrtea, operations_txt=operations_example, excel_dcf_model=None):
    styles = getSampleStyleSheet()
    custom_bodytext = ParagraphStyle(name='CustomBodyText',
                                     fontName='Times-Roman',
                                     fontSize=10,
                                     leading=14,
                                     textColor=colors.black)

    title_style = ParagraphStyle(name='TitleStyle',
                                 fontSize=32,
                                 leading=24,
                                 bold=True,
                                 textColor=colors.black,
                                 spaceAfter=20)

    # save point and file name
    doc = SimpleDocTemplate(f"reports/{date}_{save_name.lower()}.pdf")
    doc.leftMargin, doc.rightMargin = 30, 30
    doc.topMargin, doc.bottomMargin = 30, 30
    story = []
    empty_line = Spacer(1, 20)

    report_title = Paragraph(f"{name.title()} Report", title_style)
    story.append(report_title)

    report_date = Paragraph(
        f"{datetime.strptime(date, '%Y%m%d').strftime('%d.%m.%Y')}", styles["BodyText"])
    story.append(report_date)

    report_close = Paragraph(f"Stock closed at {close}$", styles["BodyText"])
    story.append(report_close)

    report_rating = Paragraph(
        f"Rating: {rating} with a target value of {target_value}$", styles["BodyText"])
    story.append(report_rating)
    story.append(empty_line)

    description = [Paragraph("About", styles["h2"]),
                   Paragraph(description, custom_bodytext)]
    story.append(description[0])
    story.append(description[1])

    operations = [Paragraph("Operations", styles["h2"]),
                  Paragraph(operations_txt, custom_bodytext)]
    story.append(operations[0])
    story.append(operations[1])

    if excel_dcf_model is not None:
        img = Image(excel_dcf_model)
        img.drawWidth = doc.width * 0.8
        img.drawHeight = img.drawWidth * img.aspectRatio
        story.append(img)

    doc.build(story)
