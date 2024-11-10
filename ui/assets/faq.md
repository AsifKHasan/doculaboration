## Prerequisites

1. Share the gsheet with [celloscope-2024@celloscope-160107.iam.gserviceaccount.com](mailto:celloscope-2024@celloscope-160107.iam.gserviceaccount.com)


## How do I add images to the gsheet?

1. Get the credential for `Sepectrum ftp server` from [asif.hasan@gmail.com](mailto:asif.hasan@gmail.com)
2. Upload the image to the spectrum ftp server using the [FileZilla Client](https://filezilla-project.org/download.php?type=client)
3. Refer the image in the gsheet
![image](resource:assets/refer_image_in_gsheet.png)
4. Add a caption row to render the image in the list of figures
![image](resource:assets/image_caption_row.png)
5. Insert the following note in the caption row
```bash
    {"style": "Figure", "keep-with-next": true}
```
![image](resource:assets/image_caption_row_note.png)


## How do I add tables to the gsheet?

1. Add a caption row to render the table in the list of tables
![image](resource:assets/table_caption_row.png)
2. Insert the following note in the caption row
```bash
    {"content": "free", "style": "Table", "keep-with-next": true}
```
![image](resource:assets/table_caption_row_note.png)

## What content width should I use in the gsheet?

Content width should be *800*
![image](resource:assets/content_width.png)

