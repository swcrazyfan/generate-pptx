import os
import openai
import json
import re
import fontawesome as fa
import tempfile
import boto3
from botocore.exceptions import NoCredentialsError
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from pptx.util import Pt
import streamlit as st

fonts_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
os.environ["FONTCONFIG_PATH"] = fonts_path

openai.api_key = os.environ["OPENAI_API_KEY"]

def apply_theme(presentation, theme):
	if theme == "dark":
		background_color = RGBColor(43, 43, 43)
		text_color = RGBColor(255, 255, 255)
		font_name = "Calibri"
	elif theme == "light":
		background_color = RGBColor(239, 239, 239)
		text_color = RGBColor(32, 32, 32)
		font_name = "Calibri"
	elif theme == "corporate":
		background_color = RGBColor(46, 117, 182)
		text_color = RGBColor(255, 255, 255)
		font_name = "Arial"
	elif theme == "playful":
		background_color = RGBColor(255, 204, 102)
		text_color = RGBColor(32, 32, 32)
		font_name = "Comic Sans MS"
	elif theme == "modern":
		background_color = RGBColor(45, 62, 80)
		text_color = RGBColor(255, 255, 255)
		font_name = "Segoe UI"
	elif theme == "vibrant":
		background_color = RGBColor(236, 98, 128)
		text_color = RGBColor(255, 255, 255)
		font_name = "Verdana"
	else:
		raise ValueError("Unsupported theme")

	for slide in presentation.slides:
		slide.background.fill.solid()
		slide.background.fill.fore_color.rgb = background_color
		for shape in slide.shapes:
			if shape.has_text_frame:
				for paragraph in shape.text_frame.paragraphs:
					for run in paragraph.runs:
						run.font.color.rgb = text_color
						run.font.name = font_name
						if theme in ["corporate", "modern", "vibrant"]:
							run.font.bold = True

	return presentation

def upload_to_s3_and_get_temporary_url(bucket_name, file_path, file_key, expiration=3600):
	aws_access_key_id = os.environ.get("AWS_ACCESS_KEY_ID")
	aws_secret_access_key = os.environ.get("AWS_SECRET_ACCESS_KEY")
	
	s3 = boto3.client(
		"s3",
		aws_access_key_id=aws_access_key_id,
		aws_secret_access_key=aws_secret_access_key,
	)
	
	try:
		s3.upload_file(file_path, bucket_name, file_key)
		response = s3.generate_presigned_url(
			"get_object",
			Params={"Bucket": bucket_name, "Key": file_key},
			ExpiresIn=expiration,
		)
		return response
	except NoCredentialsError as e:
		raise BadRequestError(f"Credentials not found: {str(e)}")

def generate_pptx(lesson_topic):
	prompt = (
		# ... (same as before)
	)

	full_prompt = prompt.format(lesson_topic=lesson_topic)

	response = openai.ChatCompletion.create(
		model="gpt-4",
		messages=[
			# ... (same as before)
		],
		max_tokens=650,
		n=1,
		stop=None,
		temperature=0.7,
	)

	output = response.choices[0].text.strip()

	theme_pattern = re.compile(r"(dark|light|corporate|playful|modern|vibrant)")
	theme_match = theme_pattern.search(output)
	theme = theme_match.group(0)

	output = theme_pattern.sub("", output).strip()

	ppt = Presentation()

	slides_data = output.split('\n\n')

	for slide_data in slides_data:
		slide_info = slide_data.split('\n')

		slide = ppt.slides.add_slide(ppt.slide_layouts[1])

		title = slide.shapes.title
		title.text = slide_info[0].split(':', 1)[-1].strip()

		content = slide_info[1:]
		content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4), Inches(4))
		content_text = content_box.text_frame

		placeholder_shape = None

		has_image_or_icon = False
		for line in content:
			if re.match(r"Image Placeholder|Icon:", line):
				image_placeholder_left = Inches(4.5)
				placeholder_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, image_placeholder_left, Inches(1.5), Inches(2), Inches(2))
				if "Icon" in line:
					icon_code = line.split("Icon:")[-1].strip()
					icon = fa.icons[icon_code]
					placeholder_shape.text_frame.text = icon
					placeholder_shape.text_frame.paragraphs[0].runs[0].font.name = "Font Awesome 5 Free"
					placeholder_shape.text_frame.paragraphs[0].runs[0].font.bold = True
					placeholder_shape.text_frame.paragraphs[0].runs[0].font.size = Pt(72)
				else:
					placeholder_shape.text_frame.text = line.strip()
					placeholder_shape.text_frame.paragraphs[0].runs[0].font.bold = True
					placeholder_shape.text_frame.paragraphs[0].runs[0].font.size = Inches(0.25)
			else:
				paragraph = content_text.add_paragraph()
				paragraph.text = line.strip()
				paragraph.space_before = Inches(0.1)

		apply_theme(ppt, theme)

		with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt_file:
			ppt.save(tmp_ppt_file.name)
			temp_file_path = tmp_ppt_file.name

			s3_bucket_name = os.environ["S3_BUCKET_NAME"]
			file_key = f"{lesson_topic.replace(' ', '_')}_presentation.pptx"
			presigned_url = upload_to_s3_and_get_temporary_url(s3_bucket_name, temp_file_path, file_key, expiration=3600)

	return {'temporary_url': presigned_url}

st.title("PowerPoint Presentation Generator")
lesson_topic = st.text_input("Enter the lesson topic:")

if st.button("Generate Presentation"):
	if lesson_topic:
		result = generate_pptx(lesson_topic)
		st.write(f"Your presentation is ready! [Download it here]({result['temporary_url']})")
	else:
		st.warning("Please enter a lesson topic.")