# Copyright (c) Leapfrog Direct Response LLC
#    (https://bitbucket.org/leapfrogdevelopment/)
#
# PunkyBrowster is a derivative work of Spynner:
#
#
#    Copyright (c) Arnau Sanchez <tokland@gmail.com>
#
#    This script is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this software.  If not, see <http://www.gnu.org/licenses/>
#
#
# PunkyBrowster was forked from the 2009-07-09 edition of Spynner.
#
# Additional Terms Pertaining Only to the Modified Aspects of this Software:
# Neither the name of Leapfrog Direct Response, LLC, including its subsidiaries
# and affiliates nor the names of its contributors, may be used to endorse or
# promote products derived from this software without specific prior written
# permission.


import sys
import time

from PyQt4.QtCore import SIGNAL, QUrl, Qt, QSize
from PyQt4.QtGui import QApplication, QImage, QPainter
from PyQt4.QtNetwork import QNetworkRequest
from PyQt4.QtWebKit import QWebElement, QWebPage, QWebView, QWebSettings

SUPPORTED_STRING_INPUT_TYPES = ('text',
                                'password',
                                'datetime',
                                'datetime-local',
                                'date',
                                'month',
                                'time',
                                'week',
                                'number',
                                'range',
                                'email',
                                'url',
                                'search',
                                'tel',
                                'color')


if sys.platform == 'darwin':
    from PyQt4.QtCore import pyqtRemoveInputHook
    pyqtRemoveInputHook()

__all__ = ['Browster']


def func(hey=[], dodo={}):
    hey.append(1)
    dodo['a'] = 3
    return 555


class Browster(object):
    application = QApplication([])
    event_looptime = 0.01

    # Feihong: Does this constructor need another keyword argument for
    # JavascriptEnabled, JavascriptCanOpenWindows?
    # See http://doc.trolltech.com/4.6/qwebsettings.html#WebAttribute-enum
    def __init__(self, auto_load_images=True, ignore_ssl_errors=False, webpage=None):
        """Init a Browser instance. 'auto_load_images' is fairly self-explanatory,
        'ignore_ssl_errors' will silence any SSL problems.

        'webpage' allows you to specify your own webpage object for use with the browser.
        By default, it uses the QWebPage object from PyQT. For details, see see:
        http://doc.qt.nokia.com/4.5/qwebpage.html

        """
        self.webpage = webpage or QWebPage()

        self.webpage.settings().setAttribute(QWebSettings.AutoLoadImages, auto_load_images)

        self.webframe = self.webpage.mainFrame()

        self.webview = None

        self._load_status = None

        self.statuses = {}

        self.ignore_ssl_errors = ignore_ssl_errors

        self.webpage.connect(
            self.webpage,
            SIGNAL('loadFinished(bool)'),
            self._on_load_finished
        )
        self.webpage.connect(
            self.webpage,
            SIGNAL("loadStarted()"),
            self._on_load_started
        )
        self.webpage.connect(
            self.webpage.networkAccessManager(),
            SIGNAL("sslErrors(QNetworkReply *,const QList<QSslError>&)"),
            self._on_ssl_error
        )
        self.webpage.connect(
            self.webpage.networkAccessManager(),
            SIGNAL("finished(QNetworkReply *)"),
            self._on_finished
        )

    def all(self, selector):
        """Return a list of all elements in the current document that match the
        given CSS selector.

        """
        result = self.webframe.documentElement().findAll(selector)

        # The QWebElement.findAll() method returns an instance of
        # QWebElementCollection, which is NOT iterable.
        return [result[i] for i in xrange(result.count())]

    def count(self, selector):
        """Return the number of elements that match the given CSS selector."""
        return self.webframe.documentElement().findAll(selector).count()

    def _first(self, selector):
        """Return the first element in the current document that matches the
        given CSS selector.

        If no element is found, raise PunkyElementNotFound exception.

        """
        el = self.webframe.documentElement().findFirst(selector)

        # The QWebElement.findFirst() method ALWAYS returns an instance of
        # QWebElement, even if no matches are found.
        if not el.isNull():
            return el
        else:
            raise PunkyElementNotFound("No page element found matching: {0}".format(selector))

    def first(self, selector):
        try:
            return self._first(selector)
        except PunkyElementNotFound:
            return None

    def _events_loop(self, wait=None):
        """Let the QApplication object process some events, then sleep."""
        if wait is None:
            wait = self.event_looptime

        self.application.processEvents()
        time.sleep(wait)

    def _on_webview_destroyed(self, window):
        self.webview = None

    def _on_load_started(self):
        self._load_status = None
        self.statuses.clear()

    def _on_load_finished(self, successful):
        self._load_status = successful

    def _on_finished(self, reply):
        url = unicode(reply.url().toString())
        status = str(reply.attribute(QNetworkRequest.HttpStatusCodeAttribute)
                     .toString())
        self.statuses[url] = int(status) if status else None

    def _on_ssl_error(self, reply, errors):
        if self.ignore_ssl_errors:
            reply.ignoreSslErrors()
        else:
            raise PunkyError('SSL error(s) encountered. (Instantiate Browster'
                             ' with ignore_ssl_errors=True to suppress'
                             ' these).\n\nErrors:\n' +
                '\n'.join(unicode(error.errorString()) for error in errors))

    def _wait_load(self, timeout=None):
        """Block until a web page is loaded.

        If no page request was initiated, immediately return the load status
        of the current page. Otherwise keep waiting until the requested page
        finishes loading or the timeout period has elapsed. If the page does
        not timeout, then return the load status.

        """
        self._events_loop(0.0)

        if self._load_status is not None:
            load_status = self._load_status
            self._load_status = None
            return load_status

        # Keep waiting for the load to finish, until timeout occurs:
        itime = time.time()
        while self._load_status is None:
            if timeout and time.time() - itime > timeout:
                raise PunkyTimeout("Timeout reached: %d seconds" % timeout)
            self._events_loop()

        self._events_loop(0.0)

        load_status = self._load_status
        self._load_status = None
        return load_status

    @property
    def doc(self):
        """Return the QWebElement instance representing the current document."""
        return self.webframe.documentElement()

    @property
    def html(self):
        """Return a unicode string representing the HTML code of the current
        frame.

        """
        return unicode(self.webframe.toHtml())

    @property
    def status(self):
        """Return the HTTP status code received for the current location."""
        try:
            return self.statuses[self.url]
        except KeyError:
            pass

    @property
    def url(self):
        """Return the URL of the current frame."""
        return unicode(self.webframe.url().toString())

    #{ Basic interaction with browser

    def load(self, url, timeout=None):
        """Initiate a page request and wait until the page either finishes
        loading or times out.

        Return the load status when done.

        """
        self.webframe.load(QUrl(url))
        return self._wait_load(timeout)

    def click(self, selector_or_element, wait_load=False, timeout=None):
        """Click any clickable element on the current page. You can pass in
        either a QWebElement instance, or string representing a CSS selector.
        If you pass in CSS, the first matching element will be the one clicked."""

        if isinstance(selector_or_element, QWebElement):
            el = selector_or_element
        else:
            el = self._first(selector_or_element)

        # Because we must create a new variable to initialize the mouse event
        # object, we must do the work inside an anonymous function in order to
        # not pollute global scope.
        el.evaluateJavaScript("""
        (function(el) {
            var evt = document.createEvent('MouseEvents');
            evt.initEvent('click', true, true);
            el.dispatchEvent(evt);
        })(this);
        """)

        if wait_load:
            return self._wait_load(timeout)

    def click_link(self, selector, timeout=None):
        """Click a link and wait for the page to load."""
        return self.click(selector, wait_load=True, timeout=timeout)

    def wait_load(self, timeout=None):
        """Wait until the page is loaded."""
        return self._wait_load(timeout)

    def wait(self, waittime):
        """Wait some time.

        This is an active wait, the events loop will be run, so it
        may be useful to wait for synchronous Javascript events that
        change the DOM.

        """
        itime = time.time()
        while time.time() - itime < waittime:
            self._events_loop()

    def close(self):
        """Close Browser instance and release resources."""
        if self.webview:
            self.destroy_webview()

        if self.webpage:
            del self.webpage

    #{ Webview

    def create_webview(self, show=False, focus=False):
        """Create a QWebView object and insert current QWebPage into it."""
        if self.webview:
            raise PunkyError("Cannot create webview (already initialized)")

        self.webview = QWebView()
        self.webview.setPage(self.webpage)

        window = self.webview.window()
        window.setWindowTitle('PunkyBrowster')   # should this be configurable?
        window.setAttribute(Qt.WA_DeleteOnClose)
        window.connect(window, SIGNAL('destroyed(QObject *)'),
                       self._on_webview_destroyed)

        if not focus:
            window.setAttribute(Qt.WA_ShowWithoutActivating)  # Stop stealing my window focus all the time.

        if show:
            self.show()

    def destroy_webview(self):
        """Destroy current QWebView."""
        if not self.webview:
            raise PunkyError("Cannot destroy webview (not initialized)")
        del self.webview

    def show(self):
        """Show webview browser."""
        if not self.webview:
            raise PunkyError("Webview is not initialized")
        self.webview.show()

    def browse(self):
        """Let the user browse the current page (infinite loop)."""
        if not self.webview:
            raise PunkyError("Webview is not initialized")

        self.show()

        while self.webview:
            self._events_loop()

    #}

    #{ Form manipulation

    def submit(self, selector, wait_load=False, timeout=None):
        el = self._first(selector)

        if el.tagName() == 'FORM':
            el.evaluateJavaScript('this.submit()')
        else:
            raise PunkyError('Cannot call submit() on a non-form element')

        if wait_load:
            return self._wait_load(timeout)

    def css(self, selector, css):
        """Modify the CSS style of an element using a CSS selector."""
        self._first(selector).setAttribute('style', css)

    def fill(self, selector, value):
        """Fill an input (of type text) or textarea with a string value using a
        CSS selector.

        NOTE: Using QWebElement.setAttribute() only works for text inputs. But
        setting this.value through JS works for both text inputs AND text areas.

        """
        el = self._first(selector)

        if el.tagName() == "TEXTAREA" or (el.tagName() == "INPUT" and
                          el.attribute('type') in SUPPORTED_STRING_INPUT_TYPES):
            # Make sure to handle double-quotes by escaping them
            el.evaluateJavaScript(
                u'this.value="{0}"'.format(value.replace('"', '\\"')))
        else:
            raise PunkyFormControlMismatch(
                "Cannot 'fill()' on {0} element {1}".format(
                    el.tagName(),
                    dict((str(i), str(el.attribute(i)))
                        for i in el.attributeNames())
                    )
                )


    def check(self, selector):
        """Check a checkbox."""
        el = self._first(selector)

        if el.tagName() == "INPUT" and el.attribute('type') in ('radio', 'checkbox'):
            el.evaluateJavaScript('this.checked=true;')
        else:
            raise PunkyFormControlMismatch("Cannot 'check()' {0} element {1}".format(
                    el.tagName(),
                    dict((str(i), str(el.attribute(i)))
                        for i in el.attributeNames())
                    )
                )

    def uncheck(self, selector):
        """Uncheck a checkbox."""
        el = self._first(selector)

        if el.tagName() == "INPUT" and el.attribute('type') == 'checkbox':
            el.evaluateJavaScript('this.checked=false;')
        else:
            raise PunkyFormControlMismatch("Cannot 'uncheck()' {0} element {1}".format(
                    el.tagName(),
                    dict((str(i), str(el.attribute(i)))
                        for i in el.attributeNames())
                    )
                )

    choose = check          # choose a radio button

    def select(self, selector):
        """Mark a select widget's option selected."""
        el = self._first(selector)

        if el.tagName() == "OPTION":
            # Make sure to mark the option as selected, but also manually
            # trigger the 'change' event of the SELECT element
            el.evaluateJavaScript("""
                this.selected = true;
                (function(selectEl) {
                    var evt = document.createEvent('HTMLEvents');
                    evt.initEvent('change', true, true);
                    selectEl.dispatchEvent(evt);
                })(this.parentNode);
            """)
        else:
            raise PunkyFormControlMismatch("Cannot 'select()' {0} element {1}".format(
                    el.tagName(),
                    dict((str(i), str(el.attribute(i)))
                        for i in el.attributeNames())
                    )
                )

    #}

    #{ Extract values

    def value(self, selector):
        """Return the value of the first element that matches the given
        selector.

        If no input[type=text] or textarea matches the selector, an
        exception will be raised.

        """
        el = self._first(selector)
        if el.tagName() == 'TEXTAREA' or \
           (el.tagName() == 'INPUT' and str(el.attribute('type')) == 'text'):
            v = self._first(selector).evaluateJavaScript('this.value')
            return unicode(v.toString())
        else:
            raise PunkyError('value() expected input[type=text] or textarea')

    def checked(self, selector):
        """If the given selector references a checkbox or radio button, return
        whether it is checked.

        Otherwise raise an exception.

        """
        el = self._first(selector)
        if el.tagName() == 'INPUT':
            input_type = str(el.attribute('type'))
            if input_type in ('radio', 'checkbox'):
                v = self._first(selector).evaluateJavaScript('this.checked')
                return v.toBool()

        raise PunkyError('checked() expected input[type=radio] or input[type=checkbox]')

    def selected(self, selector):
        """If the given selector references an option element, return whether it
        is selected.

        Otherwise raise an exception.

        """
        el = self._first(selector)
        if el.tagName() == 'OPTION':
            v = self._first(selector).evaluateJavaScript('this.selected')
            return v.toBool()

        raise PunkyError('selected() expected option')

    def text(self, selector):
        """Return all text underneath the first element that matches the given
        selector.

        If no element matches the selector, an exception will be
        raised.

        """
        return unicode(self._first(selector).toPlainText())

    #}

    #{ Javascript

    def runjs(self, jscode):
        """Evaluate some JavaScript on the current frame.

        NOTE: Any variables created inside `jscode` will end up in the global
        scope! (Unless you put them inside a function)

        """
        return self.webframe.evaluateJavaScript(jscode)

    #}

    #{ Miscellaneous

    def save_snapshot(self, filename='snapshot.png', element=None):
        """Take an image snapshot of the current frame. If a specific element is
        specified, only grab the image of that one element.
        """
        format = QImage.Format_ARGB32

        # First reset the viewport, then set it to the content size. If the
        # viewport is not reset, the resulting image might be larger than
        # necessary.
        self.webpage.setViewportSize(QSize(1, 1))
        self.webpage.setViewportSize(self.webframe.contentsSize())

        if element is None:
            image = QImage(self.webframe.contentsSize(), format)
            painter = QPainter(image)
            self.webframe.render(painter)
            painter.end()
        else:
            # If element is actually a CSS selector:
            if isinstance(element, basestring):
                element = self._first(element)

            rect = element.geometry()
            w, h = rect.width(), rect.height()

            x1 = rect.x()
            y1 = rect.y()
            x2 = x1 + w
            y2 = y1 + h

            image = QImage(QSize(x2, y2), format)
            painter = QPainter(image)
            self.webframe.render(painter)
            painter.end()
            image = image.copy(x1, y1, w, h)

        image.save(filename)

    #}


class PunkyError(Exception):
    """General Punky error."""


class PunkyElementNotFound(PunkyError):
    """No element found"""


class PunkyFormControlMismatch(PunkyError):
    """Attempted to call a form-manipulation method on a control that doesn't support it."""


class PunkyTimeout(PunkyError):
    """A timeout (usually on page load) has been reached."""
