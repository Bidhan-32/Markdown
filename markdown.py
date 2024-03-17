import os
import mammoth


docx_file_path = r"C:\Users\075bc\Desktop\markdown\Test_marker.docx"
output_md_dir = r"C:\Users\075bc\Desktop\markdown"
output_image_dir = r"C:\Users\075bc\Desktop\markdown\image"



def convert_html_to_markdown(input_html_file, output_md_file, output_image_file):
    command = f"pandoc -i {input_html_file} -o {output_md_file} --extract-media {output_image_file}"
    os.system(command)

with open(docx_file_path, "rb") as docx_file:
    result = mammoth.convert_to_html(docx_file)
    html_output = result.value 
    messages = result.messages

html_file_path = os.path.join(output_md_dir, "result.html")
with open(html_file_path, "w", encoding="utf-8") as html_file:
    html_file.write(html_output)


output_md_path = os.path.join(output_md_dir, "result.md")

convert_html_to_markdown(html_file_path, output_md_path, output_image_dir)
