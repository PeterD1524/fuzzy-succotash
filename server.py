import pathlib
import secrets
import tempfile
import traceback

import fastapi
import fastapi.responses
import pywintypes

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

    def cleanup(self):
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
    shared.cleanup()


@app.post(
    '/powerpoint/export_as_fixed_format2/',
    response_class=fastapi.responses.FileResponse
)
async def export_as_fixed_format2(
    file: fastapi.UploadFile,
    tmpdir: pathlib.Path = fastapi.Depends(TemporaryDirectory),
):
    tmp_stem = secrets.token_urlsafe()
    file_name = tmpdir / f'{tmp_stem}.in'
    file_name.write_bytes(await file.read())
    path = str(tmpdir / f'{tmp_stem}.out')

    try:
        presentation = shared.powerpoint_application.presentations.open(
            str(file_name),
            read_only=office.MsoTriState.msoTrue,
            with_window=office.MsoTriState.msoFalse
        )
    except pywintypes.com_error:
        traceback.print_exc()
        return fastapi.Response(
            status_code=fastapi.status.HTTP_422_UNPROCESSABLE_ENTITY
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

    stem = pathlib.Path(file.filename).stem
    if stem == '':
        filename = None
    else:
        filename = f'{stem}.pdf'

    return fastapi.responses.FileResponse(path, filename=filename)
