# -*- coding: utf-8 -*-
# Windows PE 版本信息（供 PyInstaller 使用）

VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=(1, 0, 0, 0),
        prodvers=(1, 0, 0, 0),
        mask=0x3f,
        flags=0x0,
        OS=0x40004,          # VOS_NT_WINDOWS32
        fileType=0x1,        # VFT_APP
        subtype=0x0,
    ),
    kids=[
        StringFileInfo([
            StringTable(
                u'080404b0',   # 0804 = 简体中文, 04b0 = Unicode
                [
                    StringStruct(u'CompanyName',      u'Hum0ro VX\uff1aXNHSDJ'),
                    StringStruct(u'FileDescription',  u'PDF转Word 智能文档转换工具 | 作者VX\uff1aXNHSDJ'),
                    StringStruct(u'FileVersion',      u'1.0.0.0'),
                    StringStruct(u'InternalName',     u'PDF2Word'),
                    StringStruct(u'LegalCopyright',   u'Copyright \u00a9 2025 Hum0ro. All rights reserved. VX\uff1aXNHSDJ'),
                    StringStruct(u'OriginalFilename', u'PDF2Word.exe'),
                    StringStruct(u'ProductName',      u'PDF转Word'),
                    StringStruct(u'ProductVersion',   u'1.0.0.0'),
                    StringStruct(u'LegalTrademarks',  u'Hum0ro & Jack'),
                ]),
        ]),
        VarFileInfo([VarStruct(u'Translation', [2052, 1200])]),
    ],
)
