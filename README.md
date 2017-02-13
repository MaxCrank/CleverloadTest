# Cleverload .NET desktop appliance test application

Application is intended to show the ability to use Open XML for MS Word processing development. It gives a user an ability to load MS Word files of old binary or new XML formats and process them in a predefined way using Open XML. Internal viewer is included, so changes are visible right away in the app.

Application requires [.NET Framework 4.0](https://www.microsoft.com/en-US/download/details.aspx?id=17851), [Open XML SDK](https://www.microsoft.com/en-us/download/details.aspx?id=30425) and [Microsoft Office] (https://www.office.com/) to be installed on at least XP version of [Windows](https://www.microsoft.com/en-US/windows/) operating system.

## To avoid confusion, some processing requirements will be clarified:

1. Personal info "row" was considered as a row of text, not a table row.
2. First "paragraph" was considered as an [Open XML Paragraph](https://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.paragraph(v=office.15).aspx), not a visual paragraph. So, personal info row will be added to the [test.doc](test.doc) after the first Open XML Paragraph, which is visually a header.
