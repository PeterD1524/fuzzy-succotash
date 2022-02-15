import dataclasses
import enum
from typing import Optional

import win32com.client

import office


class PpFixedFormatIntent(enum.IntEnum):
    '''PpFixedFormatIntent enumeration (PowerPoint)

    Constants that specify the intent of the fixed-format file export, passed to the ExportAsFixedFormat method of the Presentation object.

    https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfixedformatintent
    '''
    ppFixedFormatIntentPrint = 2  # Intent is to print exported file.
    ppFixedFormatIntentScreen = 1  # Intent is to view exported file on screen.


class PpFixedFormatType(enum.IntEnum):
    '''PpFixedFormatType enumeration (PowerPoint)

    Constants that specify the type of fixed-format file to export, passed to the ExportAsFixedFormat method of the Presentation object.

    https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfixedformattype
    '''
    ppFixedFormatTypePDF = 2  # PDF format
    ppFixedFormatTypeXPS = 1  # XPS format


class PpPrintHandoutOrder(enum.IntEnum):
    '''PpPrintHandoutOrder enumeration (PowerPoint)

    Specifies the page layout order in which slides appear on printed handouts that show multiple slides on one page.

    https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprinthandoutorder
    '''
    ppPrintHandoutHorizontalFirst = 2  # Slides are ordered horizontally, with the first slide in the upper-left corner and the second slide to the right of it. If your language setting specifies a right-to-left language, the first slide is in the upper-right corner with the second slide to the left of it.
    ppPrintHandoutVerticalFirst = 1  # Slides are ordered vertically, with the first slide in the upper-left corner and the second slide below it. If your language setting specifies a right-to-left language, the first slide is in the upper-right corner with the second slide below it.


class PpPrintOutputType(enum.IntEnum):
    '''PpPrintOutputType enumeration (PowerPoint)

    A value that indicates which component (slides, handouts, notes pages, or an outline) of the presentation is to be printed.

    https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
    '''
    ppPrintOutputBuildSlides = 7  # Build Slides
    ppPrintOutputFourSlideHandouts = 8  # Four Slide Handouts
    ppPrintOutputNineSlideHandouts = 9  # Nine Slide Handouts
    ppPrintOutputNotesPages = 5  # Notes Pages
    ppPrintOutputOneSlideHandouts = 10  # Single Slide Handouts
    ppPrintOutputOutline = 6  # Outline
    ppPrintOutputSixSlideHandouts = 4  # Six Slide Handouts
    ppPrintOutputSlides = 1  # Slides
    ppPrintOutputThreeSlideHandouts = 3  # Three Slide Handouts
    ppPrintOutputTwoSlideHandouts = 2  # Two Slide Handouts


class PpPrintRangeType(enum.IntEnum):
    '''PpPrintRangeType enumeration (PowerPoint)

    Specifies the type of print range for the presentation.

    https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintrangetype
    '''
    ppPrintAll = 1  # Print all slides in the presentation.
    ppPrintCurrent = 3  # Print the current slide from the presentation.
    ppPrintNamedSlideShow = 5  # Print a named slideshow.
    ppPrintSelection = 2  # Print a selection of slides.
    ppPrintSlideRange = 4  # Print a range of slides.


@dataclasses.dataclass
class Presentation:
    com: win32com.client.CDispatch

    def close(self):
        self.com.Close()

    def export_as_fixed_format2(
        self,
        path,
        fixed_format_type,
        intent=PpFixedFormatIntent.ppFixedFormatIntentScreen,
        frame_slides=office.MsoTriState.msoFalse,
        handout_order=PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
        output_type=PpPrintOutputType.ppPrintOutputSlides,
        print_hidden_slides=office.MsoTriState.msoFalse,
        print_range=None,
        range_type=PpPrintRangeType.ppPrintAll,
        slide_show_name=None,
        include_doc_properties=False,
        keep_irm_settings=True,
        doc_structure_tags=True,
        bitmap_missing_fonts=True,
        use_iso19005_1=False,
        include_markup=False,
        external_exporter=None
    ):
        kwargs = {
            'Path': path,
            'FixedFormatType': fixed_format_type,
            'Intent': intent,
            'FrameSlides': frame_slides,
            'HandoutOrder': handout_order,
            'OutputType': output_type,
            'PrintHiddenSlides': print_hidden_slides,
            'PrintRange': print_range,
            'RangeType': range_type,
            'SlideShowName': slide_show_name,
            'IncludeDocProperties': include_doc_properties,
            'KeepIRMSettings': keep_irm_settings,
            'DocStructureTags': doc_structure_tags,
            'BitmapMissingFonts': bitmap_missing_fonts,
            'UseISO19005_1': use_iso19005_1,
            'IncludeMarkup': include_markup,
        }
        if external_exporter is not None:
            kwargs['ExternalExporter'] = external_exporter
        return self.com.ExportAsFixedFormat2(**kwargs)


@dataclasses.dataclass
class Presentations:
    com: win32com.client.CDispatch

    def open(
        self,
        file_name,
        read_only=office.MsoTriState.msoFalse,
        untitled=office.MsoTriState.msoFalse,
        with_window=office.MsoTriState.msoTrue
    ):
        presentation = Presentation(
            self.com.Open(
                FileName=file_name,
                ReadOnly=read_only,
                Untitled=untitled,
                WithWindow=with_window
            )
        )
        return presentation


@dataclasses.dataclass
class Application:
    com: Optional[win32com.client.CDispatch] = None

    def __post_init__(self) -> None:
        if self.com is None:
            self.com = win32com.client.Dispatch('PowerPoint.Application')

    def quit(self):
        self.com.Quit()

    @property
    def presentations(self):
        return Presentations(self.com.Presentations)
