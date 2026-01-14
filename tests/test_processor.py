from capslide import *
import os.path
import pytest
import json

base_dir = os.path.dirname(os.path.abspath(__file__))
templates_dir = os.path.join(base_dir, 'templates')
outputs_dir = os.path.join(base_dir, 'outputs')
data_dir = os.path.join(base_dir, 'data')

default_template_path = os.path.join(templates_dir, 'sample.pptx')

def get_output_file_path(filename):
    return os.path.join(outputs_dir, filename)

def get_template_file_path(filename):
    return os.path.join(templates_dir, filename)

def get_data_file_path(filename):
    return os.path.join(data_dir, filename)

def test_subtitles_processor_exceptions():
    with pytest.raises(PowerPointTemplateNotFoundError):
        SubtitlesProcessor(
            output_path=get_output_file_path('dest.pptx'),
            template_path=get_template_file_path('non_existent_template.pptx'),
            template_slide_page_number=0,
            ignore_masks=False,
            verbose=True
        )
    with pytest.raises(SubtitlesTemplateSlideIndexError):
        SubtitlesProcessor(
            output_path=get_output_file_path('dest.pptx'),
            template_path=default_template_path,
            template_slide_page_number=100,
            ignore_masks=False,
            verbose=True
        )

    with pytest.raises(SubtitlesTemplateSlideIndexError):
        SubtitlesProcessor(
            output_path=get_output_file_path('dest.pptx'),
            template_path=default_template_path,
            template_slide_page_number=100,
            ignore_masks=False,
            verbose=True
        )
    with pytest.raises(SubtitlesTemplateSlideMasterError):
        SubtitlesProcessor(
            output_path=get_output_file_path('dest.pptx'),
            template_path=default_template_path,
            template_slide_page_number=1,
            ignore_masks=False,
            verbose=True
        )

    with pytest.raises(SubtitlesTemplatePlaceholderError):
        SubtitlesProcessor(
            output_path=get_output_file_path('dest.pptx'),
            template_path=default_template_path,
            template_slide_page_number=2,
            ignore_masks=False,
            verbose=True
        )
    
def test_subtitles_processor1():
    processor = SubtitlesProcessor(
        output_path=get_output_file_path('dest1.pptx'),
        template_path=default_template_path,
        template_slide_page_number=3,
        ignore_masks=False,
        verbose=True
    )
    assert processor is not None

    assert processor.get_placeholders_count(processor.template_slide, "Chinese") == 1
    assert processor.get_placeholders_count(processor.get_slide_by_page_number(processor.template_pptx, 4)) == 3
    assert processor.get_placeholders_count(processor.get_slide_by_page_number(processor.template_pptx, 5)) == 4
    
    new_slide = processor.duplicate_slide(processor.template_slide)
    assert new_slide is not None
    placeholder = "Chinese"
    text = "测试文本"
    assert processor.get_placeholders_count(new_slide, placeholder) == 1

    count = processor.replace_placeholder_of_slide(new_slide, placeholder, text)
    assert count == 1
    #processor.save()
    assert processor.get_placeholders_count(new_slide, placeholder) == 0

    processor.save()

def test_subtitles_processor_row():
    processor = SubtitlesProcessor(
        output_path=get_output_file_path('dest2.pptx'),
        template_path=default_template_path,
        template_slide_page_number=4,
        ignore_masks=False,
        verbose=True
    )
    assert processor is not None

    row = {'A': '这是字幕1', 'B': '这是字幕2', 'C': '这是字幕3'}

    assert processor.append_slide_with_row(row) == 3

    
    
    #assert processor.get_placeholders_count(new_slide, placeholder) == 0

    processor.save()

def test_subtitles_processor_rows():
    processor = SubtitlesProcessor(
        output_path=get_output_file_path('dest3.pptx'),
        template_path=default_template_path,
        template_slide_page_number=7,
        ignore_masks=False,
        verbose=True
    )
    assert processor is not None

    rows = [
        {'A': '这是字幕1', 'B': '这是字幕2', 'C': '这是字幕3'},
        {'A': '这是字幕4', 'B': '这是字幕5', 'C': '这是字幕6'},
        {'A': '这是字幕7', 'B': '这是字幕8', 'C': '这是字幕9'},
    ]

    matched_count, slides_count = processor.append_slides_with_rows(rows)
    assert matched_count == 15
    assert slides_count == 3

    
    
    #assert processor.get_placeholders_count(new_slide, placeholder) == 0

    processor.save()


def test_subtitles_processor_json_file():
    processor = SubtitlesProcessor(
        output_path=get_output_file_path('dest4.pptx'),
        template_path=default_template_path,
    )
    assert processor is not None

    json_path = get_data_file_path('sample1.json')
    
    matched_count, slides_count =  processor.append_slides_from_json_file(json_path)
    processor.save()
    assert matched_count == 15
    assert slides_count == 3    

    
    
    #assert processor.get_placeholders_count(new_slide, placeholder) == 0

    

def test_subtitles_processor_text_file():
    processor = SubtitlesProcessor(
        output_path=get_output_file_path('dest5.pptx'),
        template_path=default_template_path,
        template_slide_page_number=6,
        placeholder='subtitle',
    )
    assert processor is not None

    file_path = get_data_file_path('sample2.txt')
    
    matched_count, slides_count =  processor.append_slides_from_text_file(file_path)
    processor.save()
    assert matched_count == 6
    assert slides_count == 3    

    
def test_subtitles_processor_file1():
    processor = SubtitlesProcessor(
        output_path=get_output_file_path('dest6.pptx'),
        template_path=default_template_path,
        template_slide_page_number=6,
        placeholder='subtitle',
    )
    assert processor is not None

    file_path = get_data_file_path('sample2.txt')
    
    matched_count, slides_count =  processor.append_slides_from_file(file_path)
    processor.save()
    assert matched_count == 6
    assert slides_count == 3    


def test_subtitles_processor_file2():
    processor = SubtitlesProcessor(
        output_path=get_output_file_path('dest7.pptx'),
        template_path=default_template_path,
        template_slide_page_number=7,
        placeholder='subtitle',
    )
    assert processor is not None

    file_path = get_data_file_path('sample1.json')
    
    matched_count, slides_count =  processor.append_slides_from_file(file_path)
    processor.save()
    assert matched_count == 15
    assert slides_count == 3    

