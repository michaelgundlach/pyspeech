from setuptools import setup

setup(name='speech',
      version='0.3.5',
      py_modules=['speech'],
      #install_requires=['win32com','pythoncom'],

      description="A clean interface to Windows speech recognition " \
        "and text-to-speech capabilities.",

      long_description="""
          Allows your Windows python program to:\n
            * execute a callback when certain phrases are heard\n
            * execute a callback when any understandable text is heard\n
            * have different callbacks for different groups of phrases\n
            * convert text to speech.\n
          \n
          For this to work, you must first install pywin32 (download and
          run the appropriate version from http://tinyurl.com/5jhg29 ) and
          the Microsoft Speech kit (download and run "SpeechSDK51.exe" from
          http://tinyurl.com/5m6v2 ).
          \n
          Then you can just "import speech" and be on your way!\n
          \n
          Please let me know if you like or use this module - it would make
          my day!
          """,

      author='Michael Gundlach',
      author_email='gundlach@gmail.com',
      url='http://code.google.com/p/pyspeech/',
      keywords = "speech recognition text-to-speech text to speech",

      classifiers=[
          'Development Status :: 5 - Production/Stable',
          'Environment :: Win32 (MS Windows)',
          'Intended Audience :: Developers',
          'License :: OSI Approved :: Apache Software License',
          'Operating System :: Microsoft :: Windows',
          'Programming Language :: Python',
          'Topic :: Multimedia :: Sound/Audio :: Speech',
          'Topic :: Home Automation',
          'Topic :: Scientific/Engineering :: Human Machine Interfaces',
          'Topic :: Software Development :: Libraries :: Python Modules',
          'Topic :: Desktop Environment',
          ]

     )
