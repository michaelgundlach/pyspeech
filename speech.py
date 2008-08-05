"""
speech recognition and voice synthesis module.

Please let me know if you like or use this module -- it would make my day!

speech.py: Copyright 2008 Michael Gundlach  (gundlach at gmail)
License: Apache 2.0 (http://www.apache.org/licenses/LICENSE-2.0)

For this module to work, you must install the Microsoft Speech kit:
download and run "SpeechSDK51.exe" from http://tinyurl.com/5m6v2

Very simple usage example:

import speech

speech.say("Hello")

def L1callback(phrase, listener):
    print phrase

def L2callback(phrase, listener):
    if phrase == "wow":
        speech.stoplistening(listener)
    speech.say(phrase)

L1 = speech.listenfor(["hello", "good bye"], L1callback)
L2 = speech.listenforanything(L2callback)

# callbacks are executed on a separate events thread.

assert speech.islistening()
assert speech.islistening(L2)

speech.stoplistening(L1)
assert not speech.islistening(L1)

speech.stoplistening()
"""


"""
Listener object:
    has phrases, callback
    has context, grammar
    init gets other thread to make an event handler for him, handing the
      thread his context and listener.  adds himself to _listeners.
    islistening() should return true even if the handler hasnt been created
      because stoplistening will correctly remove listener from list and
      kill his grammar ref, so the eventhandler will immediately be dead.
      i should put a __del__ on the eventhandler to print when it dies.
      wont have a ref to it and its async (or must it be sync?)
    stoplistening clears grammar, finds his handler and calls close on it
"""

class Listener2(object):
    _all = set()

    def __init__(self, context, grammar, callback):
        self._grammar = grammar
        Listener._all.add(self)

        # Tell event thread to create an event handler to call our callback
        # upon hearing speech events
        _handlerqueue.append((context, self, callback))
        _ensure_event_thread()

    def islistening():
        """True if this listener is listening for speech."""
        return self in Listener._all

    def stoplistening():
        """Stop listening for speech."""
        try:
            Listener._all.remove(self)
        except:
            pass

        # This removes all refs to _grammar.rules so the event handler can die
        self._grammar = None




from win32com.client import constants as _constants
import win32com.client
import pythoncom
import time
import thread

# Make sure that we've got our COM wrappers generated.
from win32com.client import gencache
gencache.EnsureModule('{C866CA3A-32F7-11D2-9602-00C04F8EE628}', 0, 5, 0)

_voice = win32com.client.Dispatch("SAPI.SpVoice")
_recognizer = win32com.client.Dispatch("SAPI.SpSharedRecognizer")
_listeners = []
_handlerqueue = []
_eventthread=None

class Listener(object):
    """Returned by speech.listenfor(), to pass to speech.stoplistening()."""
    def __init__(self, callback, grammar):
        self._callback = callback
        self._grammar = grammar

_ListenerBase = win32com.client.getevents("SAPI.SpSharedRecoContext")
class _ListenerCallback(_ListenerBase):
    """Created to fire events upon speech recognition.  self._listener is
    cleared when we wish to stop handling events -- this causes us to
    lose a reference to _listener._grammar.rules, which makes this event
    handler go away. TODO: we may need to call self.close() to release the
    COM object, and we should probably make goaway() a method of self
    instead of letting people do it for us.
    """
    def __init__(self, oobj, listener, callback):
        _ListenerBase.__init__(self, oobj)
        self._listener = listener
        self._callback = callback

    def OnRecognition(self, _1, _2, _3, Result):
        if self._callback and self._listener:
            newResult = win32com.client.Dispatch(Result)
            phrase = newResult.PhraseInfo.GetText()
            self._callback(phrase, self._listener)

def say(phrase):
    """Say the given phrase out loud."""
    _voice.Speak(phrase)

def listenforanything(callback):
    """When anything resembling English is heard,
    callback(spoken_text, listener) is executed.  Execution takes
    place on a single thread shared by all listener callbacks.  Returns
    an object that can be passed to stoplistening() to stop listening,
    which is also passed as the second argument to callback.
    """
    return _startlistening(None, callback)

def listenfor(phraselist, callback):
    """If any of the phrases in the given list are heard,
    callback(spoken_text, listener) is executed.  Execution takes
    place on a single thread shared by all listener callbacks.  Returns
    an object that can be passed to stoplistening() to stop listening,
    which is also passed as the second argument to callback.
    """
    return _startlistening(phraselist, callback)

def _startlistening(phraselist, callback):
    """Starts listening in Command-and-Control mode if phraselist is
    not None, or dictation mode if phraselist is None.  When a
    phrase is heard, callback(phrase_text, listener) is executed.
    Returns an object that can be passed to stoplistening() to
    stop listening, which is also passed as the second argument to
    callback.  Ensures that a separate thread exists checking for
    speech events.
    """
    # Make a command-and-control grammar        
    context = _recognizer.CreateRecoContext()
    grammar = context.CreateGrammar()

    if phraselist:
        grammar.DictationSetState(0)
        # dunno why we pass the constants that we do here
        rule = grammar.Rules.Add("rule",
                _constants.SRATopLevel + _constants.SRADynamic, 0)
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

    # Add a request to create an event handler that's called when recognition 
    # occurs, executing callback(phrase_text, listener) on the events thread.
    _handlerqueue.append( (context, listener, callback) )
    _ensure_event_thread()

    return listener

def _ensure_event_thread():
    """
    Make sure the eventthread is running.  It checks the handlerqueue
    for new eventhandlers to create, and runs the message pump.
    """
    global _eventthread
    if not _eventthread:
        def loop():
            print "looping"
            while _eventthread:
                pythoncom.PumpWaitingMessages()
                if _handlerqueue:
                    (context,listener,callback) = _handlerqueue.pop()
                    # Just creating a _ListenerCallback object makes events
                    # fire till listener loses reference to its grammar object
                    _ListenerCallback(context, listener, callback)
                time.sleep(.5)
            print "stopping looping"
        _eventthread = 1 # so loop doesn't terminate immediately
        _eventthread = thread.start_new_thread(loop, ())

def stoplistening(listener = None):
    """
    Stop listening to the given listener.  If no listener is
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
        global _eventthread
        _eventthread = None # kill the spinner if it exists

    return stoppedSomeone

def islistening(listener = None):
    """
    True if speech is listening to the given listener, or to
    any listener if none is provided.
    """
    if not listener:
        return _listeners != []
    else:
        return listener in _listeners

