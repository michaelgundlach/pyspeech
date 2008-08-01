from win32com.client import constants
import win32com.client
import pythoncom
import time
import thread


"""
Please let me know if you like or use this module -- it would make my day!

speech.py: Copyright 2008 Michael Gundlach (gundlach at gmail)
License: Apache 2.0 (http://www.apache.org/licenses/LICENSE-2.0)

For this module to work, you must first:
1. Install the Microsoft Speech kit: download and run "SpeechSDK51.exe"
   from http://tinyurl.com/5m6v2
2. Then open PythonWin (installable via http://tinyurl.com/5ezco9) and
   choose Tools | COM MakePY utility | Microsoft Speech Object Library 5.0.

Very simple usage example:

import speech

speech.say("Hello")

L1 = speech.listenfor(["hello", "goodbye"], lambda phrase, L: print phrase)
L2 = speech.listenforanything(
       lambda txt, listener: if txt == "wow": speech.stoplistening(listener))

while speech.islistening(L2):
  speech.pump_waiting_messages() # each call sleeps .5 secs, so spinloop is OK

speech.stoplistening(L1)
"""

_loopthread = None
_listeners = []
_voice = win32com.client.Dispatch("SAPI.SpVoice")

class Listener(object):
    """Returned by speech.listenfor(), to be passed to speech.stoplistening().
    """
    def __init__(self, callback, grammar):
        self._callback = callback
        self._grammar = grammar

class _ListenerCallback(win32com.client.getevents("SAPI.SpSharedRecoContext")):
    """Created to fire events upon speech recognition.  There's no way to turn
    it off once it's been created, and there's no way (that I know of) to let
    it have an __init__ method.  So self._callback is assigned by the
    creator, and cleared when we wish to "stop" handling events (though
    self.OnRecognition will be a no-op, though it will still fire.)
    """
    def OnRecognition(self, StreamNumber, StreamPosition, RecognitionType, Result):
        if self._callback:
            newResult = win32com.client.Dispatch(Result)
            phrase = newResult.PhraseInfo.GetText()
            self._callback(phrase, self._listener)
   
def say(phrase):
    """Say the given phrase out loud.
    """
    _voice.Speak(phrase)

def listenforanything(callback):
    """When anything resembling English is heard,
    callback(spoken_text, listener) is executed.  Returns an
    object that can be passed to stoplistening() to stop listening,
    which is also passed as the second argument to callback.
    """
    return _startlistening(None, callback)
    
def listenfor(phraselist, callback):
    """If any of the phrases in the given list are heard,
    callback(spoken_text, listener) is executed.  Returns an
    object that can be passed to stoplistening() to stop listening,
    which is also passed as the second argument to callback.
    """
    return _startlistening(phraselist, callback)

_recognizer = None # TODO temp to fix problem
def _startlistening(phraselist, callback):
    """Starts listening in Command-and-Control mode if phraselist is
    not None, or dictation mode if phraselist is None.  When a
    phrase is heard, callback(phrase_text, listener) is executed.
    Returns an object that can be passed to stoplistening() to
    stop listening, which is also passed as the second argument to
    callback.
    """
    # Make a command-and-control grammar        
    global _recognizer
    if not _recognizer:
        _recognizer = win32com.client.Dispatch("SAPI.SpSharedRecognizer")
    context = _recognizer.CreateRecoContext()
    grammar = context.CreateGrammar()
    
    if phraselist:
        grammar.DictationSetState(0)
        # dunno why we pass the constants that we do here
        rule = grammar.Rules.Add("rule",
                constants.SRATopLevel + constants.SRADynamic, 0)
        rule.Clear()
    
        for phrase in phraselist:
            rule.InitialState.AddWordTransition(None, phrase)

        # not sure if this is needed - was here before but dupe is below
        grammar.Rules.Commit()
    
        # Commit the changes to the grammar
        grammar.CmdSetRuleState("rule", 1) # active
        grammar.Rules.Commit()
    else:
        grammar.DictationSetState(1)
    
    listener = Listener(callback, grammar)
    _listeners.append(listener)

    # And add an event handler that's called when recognition occurs,
    # executing callback(phrase_text, listener).
    eventHandler = _ListenerCallback(context)

    # I can't figure out how to make _ListenerCallback allow an __init__
    # method, so I've got to hook on the callback and listener here.
    eventHandler._listener = listener
    eventHandler._callback = callback
    
    return listener

def stoplistening(listener = None):
    """Stop listening to the given listener.  If no listener is
    specified, stop listening to all listeners.  Returns True if
    at least one listener existed to stop.
    """

    # Removing a listener's reference to _grammar causes us to
    # lose reference to the rule, which causes the event to go
    # away and stop firing.  Removing a listener's reference
    # to _callback is just an extra safeguard so that even if
    # the event *does* fire, nothing will happen.
    def removeListener(listener):
        listener._grammar, listener._callback = None, None
        if listener in _listeners:
            _listeners.remove(listener)

    stoppedSomeone = False
    
    if listener:
        stoppedSomeone = (listener in _listeners)
        removeListener(listener)
    else:
        stoppedSomeone = (_listeners != [])
        while _listeners:
            removeListener(_listeners[0])

    if not _listeners:
        global _loopthread
        _loopthread = None # kill the spinner if it exists

    return stoppedSomeone

def keeplistening():
    """Ensure that a thread is calling pump_waiting_messages() every
    second or so. Only one thread is created even if there are multiple
    calls.  The thread is killed when no listeners exist (when
    stoplistening() has been called on all of them or without an argument.)
    """
    global _loopthread
    if not _loopthread:
        def loop():
            print "looping"
            while _loopthread:
                pump_waiting_messages()
            print "stopping looping"

        _loopthread = 1 # so the loop code doesn't see None on startup
        _loopthread = thread.start_new_thread(loop, tuple([]))

def islistening(listener = None):
    """True if speech is listening to the given listener, or to
    any listener if none is provided.
    """
    if not listener:
        return _listeners != []
    else:
        return listener in _listeners

def pump_waiting_messages():
    """Receive all speech events in the COM queue.  Without calling this,
    events may back up and the listeners may never get their callbacks called.
    This then sleeps for .5 seconds so you can safely call it in a loop.
    """
    pythoncom.PumpWaitingMessages()
    time.sleep(.5) # so users in a spinwait don't lock the CPU
