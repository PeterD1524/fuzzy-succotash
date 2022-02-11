import enum


class MsoTriState(enum.IntEnum):
    '''MsoTriState enumeration (Office)

    Specifies a tri-state value.

    https://docs.microsoft.com/en-us/office/vba/api/office.msotristate
    '''
    msoCTrue = 1  # Not supported
    msoFalse = 0  # False
    msoTriStateMixed = -2  # Not supported
    msoTriStateToggle = -3  # Not supported
    msoTrue = -1  # True
