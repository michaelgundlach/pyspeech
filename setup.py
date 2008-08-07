from setuptools import setup

setup(name='speech',
      version='0.4.3',
      py_modules=['speech'],
      #install_requires=['win32com','pythoncom'],

      description="A clean interface to Windows speech recognition " \
        "and text-to-speech capabilities.",

      long_description="""
------------
speech.py
------------

  Allows your Windows python program to:
    * execute a callback when certain phrases are heard
    * execute a callback when any understandable text is heard
    * have different callbacks for different groups of phrases
    * convert text to speech.

Example
=======

  Showing speaking out loud and listening for all recognizable words.
  ::

    import speech
    import time

    def callback(phrase, listener):
        if phrase == "goodbye":
            listener.stoplistening()
        speech.say(phrase)

    listener = speech.listenforanything(callback)
    while listener.islistening():
        time.sleep(.5)

Requirements
============

  Requires Windows XP and Python 2.5.

  In addition to easy_installing speech.py, you'll need pywin32 (installer
  `here <http://tinyurl.com/5ezco9>`__) and the Microsoft Speech kit
  (installer `here <http://tinyurl.com/zflb>`__).

Resources
=========

  * Homepage: http://pyspeech.googlecode.com/
  * Source:

    - Browse at http://code.google.com/p/pyspeech/source/browse/trunk/

    - Get with **svn co http://pyspeech.googlecode.com/svn/trunk/
      pyspeech-read-only**

  Please let me know if you like or use this module - it would make
  my day!
""",

      author='Michael Gundlach',
      author_email='gundlach@gmail.com',
      url='http://code.google.com/p/pyspeech/',
      keywords = "speech recognition text-to-speech text to speech tts "
          "voice recognition",

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
