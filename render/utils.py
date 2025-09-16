from __future__ import annotations
from pptx.dml.color import RGBColor
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN


def hyperlink_for_rfc(rfc: str) -> str:
    return f"https://outils.change.fr/change={str(rfc).lower()}"


COLOR_MAP = {
    "urgent": RGBColor(255, 140, 0),   # orange
    "normal": RGBColor(0, 102, 204),   # blue
    "agile":  RGBColor(0, 153, 0),     # green
}
DEFAULT_COLOR = RGBColor(100, 100, 100)


def set_title(prs, slide, text: str, font_size_pt: int = 24) -> None:
    """Set a slide title into the layout's title placeholder if present.
    If no title placeholder, add a textbox at the top with a title style.
    """
    try:
        title_shape = slide.shapes.title
    except Exception:
        title_shape = None

    if title_shape is not None:
        try:
            tf = title_shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = text
            run.font.size = Pt(font_size_pt)
            run.font.bold = True
            return
        except Exception:
            pass

    # No title placeholder; add a textbox at the top
    left = Cm(1.0)
    top = Cm(1.0)
    width = prs.slide_width - Cm(2.0)
    height = Cm(1.5)
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size_pt)
    run.font.bold = True
