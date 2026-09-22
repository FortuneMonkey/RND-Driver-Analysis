import os

import win32com.client as win32

olMailItem = 0
# The MAPI property tag Outlook uses to link an attachment to a `cid:` in HTML.
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"


def build_placeholder_body(period_label, screenshot_cid):
    """
    Throwaway body content for testing -- replace with real wording whenever
    you're ready. The screenshot is referenced via cid so it renders inline.
    """
    return f"""
    <html><body style="font-family:Segoe UI, Arial, sans-serif; font-size:13px; color:#1F2328;">
        <p>Dear all,</p>
        <p>This is the report for fleet reconciliation run for
        <b>{period_label}</b>. </p>
        <p>Please find attached the latest full report and per-vehicle PDFs for your review.</p>
        <p><img src="cid:{screenshot_cid}" style="width:200px; border:1px solid #ccc;"></p>
        <p>Thank you.</p>
    </body></html>
    """


def send_report_email(
    to_recipients,
    subject,
    screenshot_path,
    attachments=None,
    cc_recipients=None,
    body_html=None,
    period_label="",
    send=False,
):
    """
    to_recipients / cc_recipients: str or list[str] of email addresses.
    screenshot_path: PNG to embed inline in the body (page 1 of the fleet PDF).
    attachments: list of file paths to attach as-is (e.g. the fleet PDF and
                 each vehicle's PDF).
    body_html: pass your own HTML to override the placeholder body.
    send: False (default) opens a draft in Outlook for review; True sends
          immediately without opening a window.
    """
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(olMailItem)

    mail.To = "; ".join(to_recipients) if isinstance(to_recipients, (list, tuple)) else to_recipients
    if cc_recipients:
        mail.CC = "; ".join(cc_recipients) if isinstance(cc_recipients, (list, tuple)) else cc_recipients
    mail.Subject = subject

    screenshot_cid = "fleet_report_page1"
    mail.HTMLBody = body_html or build_placeholder_body(period_label, screenshot_cid)

    if screenshot_path and os.path.exists(screenshot_path):
        img_attachment = mail.Attachments.Add(screenshot_path)
        img_attachment.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, screenshot_cid)

    for path in attachments or []:
        if os.path.exists(path):
            mail.Attachments.Add(path)

    if send:
        mail.Send()
    else:
        mail.Display()

    return mail
