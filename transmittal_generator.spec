# -*- mode: python -*-

block_cipher = None

added_files = [
         ( 'Docs', 'Docs' ),
         ( 'media', 'media' ),
         ( 'job_defaults', 'job_defaults' ),
         ( 'templates', 'templates' ),
         ( 'Transmittal Generated', 'Transmittal Generated' ),
         ]
         
a = Analysis(['transmittal_generator.py'],
             pathex=['C:\\Users\\IDS\\Desktop\\Under Development\\trans_gen'],
             binaries=[],
             datas=added_files,
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='transmittal_generator',
          debug=False,
          strip=False,
          upx=True,
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='transmittal_generator')
