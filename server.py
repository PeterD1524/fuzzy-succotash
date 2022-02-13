import pathlib
import tempfile

import fastapi
import fastapi.responses

import office
import powerpoint


class Shared:

    def __init__(self) -> None:
        self._powerpoint_application = None

    @property
    def powerpoint_application(self):
        if self._powerpoint_application is None:
            self._powerpoint_application = powerpoint.Application()
        return self._powerpoint_application

    def clean(self):
        if self._powerpoint_application is not None:
            self._powerpoint_application.quit()
            self._powerpoint_application = None


def TemporaryDirectory():
    with tempfile.TemporaryDirectory(dir=pathlib.Path.cwd()) as tmpdir:
        yield pathlib.Path(tmpdir)


app = fastapi.FastAPI()

shared = Shared()


@app.on_event("shutdown")
def shutdown_event():
    shared.clean()


@app.post(
    '/powerpoint/export_as_fixed_format2/',
    response_class=fastapi.responses.FileResponse
)
async def export_as_fixed_format2(
    file: fastapi.UploadFile,
    tmpdir: pathlib.Path = fastapi.Depends(TemporaryDirectory)
):
    file_name = tmpdir / file.filename
    file_name.write_bytes(await file.read())
    filename = f'{file_name.stem}.pdf'
    path = str(tmpdir / filename)
    presentation = shared.powerpoint_application.presentations.open(
        str(file_name),
        read_only=office.MsoTriState.msoTrue,
        with_window=office.MsoTriState.msoFalse
    )
    try:
        presentation.export_as_fixed_format2(
            path,
            powerpoint.PpFixedFormatType.ppFixedFormatTypePDF,
            keep_irm_settings=False,
            doc_structure_tags=False,
            bitmap_missing_fonts=False
        )
    finally:
        presentation.close()
    return fastapi.responses.FileResponse(path, filename=filename)
