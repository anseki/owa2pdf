/*
 * owa2pdf
 * https://github.com/anseki/owa2pdf
 *
 * Copyright (c) 2013 anseki
 * Licensed under the MIT license.
 */

(function() {
  'use strict';

  var page = require('webpage').create(),
    args = require('system').args,
    fs = require('fs'),
    arg, msUser, msPw, srcFile, destFile,
    jqLoad = true, ignoreFail = false, pdfSaved = false,
    uriFiles, uriAppFrame, uriApp, appClass,
    actionLoopInterval = 1000, countStep = 1, onReadyCallback, timerId, i, iLen,
    debugScrnFile, debugScrnCnt,

    TIMEOUT = 30,
    JQ_PATH = './jquery-2.0.3.min.js',
    defaultReqHeaders = {'Accept-Language': 'en-US,en;q=0.8'},
    DEBUG = 0, // >=1: Message, bit2: Screenshot, bit3: Resource I/O

    // Temporary file Utility (path, baseName, baseNameRe)
    // SkyDrive is a stickler for file name.
    tmpFile = {
      init: function(dirPath, ext) {
        this.dirPath = dirPath;
        this.ext = ext;
        this.baseNum = (new Date()).getTime();
        return this.reNew();
      },
      reNew: function() {
        var newBaseName, newPath, extStr = this.ext ? '.' + this.ext : '', i;
        for (i = 0; i < 1000; i++) {
          newBaseName = '_' + (this.baseNum++).toString(16);
          newPath = this.dirPath + fs.separator + newBaseName + extStr;
          if (!fs.exists(newPath)) {
            this.path = newPath;
            this.baseName = newBaseName + extStr;
            this.baseNameRe = new RegExp('^' +
              newBaseName.replace(/(\W)/g, '\\$1') +
                (extStr ? '(?:' + extStr.replace(/(\W)/g, '\\$1') + ')?' : '') +
              '$', 'i');
            return true;
          }
        }
        return false;
      }
    },

    selector = {
      loginUser: 'i0116',
      loginPw: 'i0118',
      loginPersistent: 'idChkBx_PWD_KMSI0Pwd',
      loginBtn: 'idSIButton9',
      logoutBtn: 'c_signout',
      uploadElm: '.cb_uploadInput input[type="file"]',
      fileItem: '.c-ListView .tileContent.officeDoc .namePlate .title,' +
              '.c-ListView .c-SetItemRow a.file .cellText',
      cmdView: '.c_mcp .uxfa_m li a',
      frameApp: 'sdx_ow_iframe',
      cmdFile: '#jewelcontainer .cui-jewel-jewelMenuLauncher',
      cmdPrint: {
        word:   '#faPrint-Menu32',
        excel:  '#m_excelWebRenderer_ewaCtl_faPrint-Menu32',
        ppoint: '#PptJewel\\.Print-Menu32'
      },
      cmdPrintSub: {
        word:   '#jbtnPrintToPdf-Menu48',
        excel:  '#m_excelWebRenderer_ewaCtl_Jewel\\.Print-Menu48',
        ppoint: '#PptJewel\\.Print\\.PrintToPdf-Menu48'
      },
      cmdExcel: {
        dialog: '#ewaDialogInner', // visibility: visible;
        optEntire: '#printEntireItem',
        btnPrint: 'button[type="submit"].ewa-dlg-button',
        ctrlArea: '#print_bar,.spacer.noprint'
      },
      linkPdf: '#PrintPDFLink',
      cmdRemove: '.c_mcp .uxfa_m li a'
    };

  for (i = 1, iLen = args.length; i < iLen; i++) {
    arg = args[i].trim();
    if      (arg === '-u') { msUser   = args[++i].trim(); }
    else if (arg === '-p') { msPw     = args[++i].trim(); }
    else if (arg === '-i') { srcFile  = args[++i].trim(); }
    else if (arg === '-o') { destFile = args[++i].trim(); }
    // else if (arg === '-s') { silent = true; }
    else if (arg === '-debug') { DEBUG = parseInt(args[++i].trim(), 10); }
  }
  if (!msUser || !msPw || !srcFile || !destFile) {
    console.error('Usage: phantomjs owa2pdf.js' +
      ' -u "user@example.com" -p "password" -i "/path/source.docx" -o "/path/dest.pdf"');
    phantom.exit(1);
  }

  // I/O files.
  if (!fs.exists(srcFile) || !fs.isFile(srcFile)) {
    console.error('-i not found: "' + srcFile + '"');
    phantom.exit(1);
  }
  (function() {
    var pos = destFile.lastIndexOf(fs.separator),
      destDirPath = pos > -1 ? destFile.substr(0, pos) : '.', // '/file' => ''
      srcBaseName, ext;
    if (!(pos > -1 ? destFile.substr(pos + 1) : destFile)) { // Check destBaseName
      console.error('-o invalid: "' + destFile + '"');
      phantom.exit(1);
    }
    if (!fs.exists(destFile)) {
      // Root directory exists always.
      // destDirPath may be root directory even if it isn't ''.
      // But check destDirPath for syntax.
      if (destDirPath) {
        if (!fs.exists(destDirPath)) {
          try {
            if (!fs.makeTree(destDirPath)) {
              console.error('Can\'t make directory: "' + destDirPath + '"');
              phantom.exit(1);
            }
          } catch(e) {
            console.error('Error(' + e.message + ') ' + e.description + ' - ' + e);
            phantom.exit(1);
          }
        } else if (!fs.isDirectory(destDirPath)) {
          console.error('Already exists that isn\'t directory: "' + destDirPath + '"');
          phantom.exit(1);
        }
      }
    } else if (!fs.isFile(destFile)) {
      console.error('-o is not file: "' + destFile + '"');
      phantom.exit(1);
    } else if (!fs.isWritable(destFile)) {
      console.error('-o not writable: "' + destFile + '"');
      phantom.exit(1);
    }
    // For temporary file.
    if (!fs.isWritable(destDirPath || '/')) {
      console.error('Directory in -o not writable: "' + destDirPath + '"');
      phantom.exit(1);
    }
    // Get ext from srcFile.
    pos = srcFile.lastIndexOf(fs.separator);
    srcBaseName = pos > -1 ? srcFile.substr(pos + 1) : srcFile;
    pos = srcBaseName.lastIndexOf('.');
    ext = pos > -1 ? srcBaseName.substr(pos + 1) : '';
    tmpFile.init(destDirPath, ext);
  })();

  // For debug mode.
  if (DEBUG) {
    if (DEBUG & 2) {
      page.viewportSize = { width: 800, height: 600 }; // For screenshot
      debugScrnFile = Math.floor((new Date()).getTime() / 1000);
      debugScrnCnt = 0;
    }
    page.onConsoleMessage = function(msg) {
      // This error occurs when page.settings.webSecurityEnabled isn't set to false
      if (msg.indexOf('Unsafe JavaScript attempt to access frame with URL') > -1)
        { return; }
      debugLog('(onConsoleMessage) ' + msg);
    };
    page.onError = function(msg, trace) {
      var msgs = [msg];
      if (trace && trace.length) {
        trace.forEach(function(t) {
          msgs.push(' -> ' + t.file + ': ' + t.line +
            (t.function ? ' (in function "' + t.function + '")' : ''));
        });
      }
      debugLog('======== onError');
      debugLog(msgs.join('\n'));
      debugLog('======== /onError');
    };
    if (DEBUG & 4) {
      page.onResourceReceived = function(response) {
        var head;
        // if ((response.url || '').indexOf('storage.live.com') < 0) { return; }
        response.url = response.url.replace(/([;:]base64,[^;]{3})[^;]+/i, '$1...');
        head = 'onResourceReceived (#' + response.id +
          ', stage "' + response.stage + '", url: ' + response.url + ')';
        debugLog('======== ' + head);
        debugLog(JSON.stringify(response, null, '  '));
        debugLog('======== /' + head);
      };
    }
  }
  // Regardless debug mode. (customHeaders)
  page.onResourceRequested = function(requestData) {
    var head;
    page.customHeaders = defaultReqHeaders; // For XMLHttpRequest.setRequestHeader
    if (DEBUG & 4) {
      // if ((requestData.url || '').indexOf('storage.live.com') < 0) { return; }
      requestData.url = requestData.url.replace(/([;:]base64,[^;]{3})[^;]+/i, '$1...');
      head = 'onResourceRequested (#' + requestData.id +
        ', url: ' + requestData.url + ')';
      debugLog('======== ' + head);
      debugLog(JSON.stringify(requestData, null, '  '));
      debugLog('======== /' + head);
    }
  };
  // Regardless debug mode. (SSL handshake failed)
  page.onResourceError = function(resourceError) {
    var head;
    if (DEBUG & 4) {
      // if ((resourceError.url || '').indexOf('storage.live.com') < 0) { return; }
      resourceError.url = resourceError.url.replace(/([;:]base64,[^;]{3})[^;]+/i, '$1...');
      head = 'onResourceError (#' + resourceError.id +
        ', url: ' + resourceError.url + ')';
      debugLog('======== ' + head);
      debugLog(JSON.stringify(resourceError, null, '  '));
      debugLog('======== /' + head);
    }
    if (resourceError.errorCode === 6) { // SSL handshake failed
      console.error(resourceError.errorString + ': ' + resourceError.url);
      console.error('"--ignore-ssl-errors=true" for phantomjs may be needed.');
      phantom.exit(1);
    }
  };
  function debugLog(msg, getScrn) {
    var filename, date, timeStamp;
    if (!DEBUG) { return; }
    date = new Date();
    timeStamp = ('00' + date.getHours()).slice(-2) +
      ':' + ('00' + date.getMinutes()).slice(-2) +
      ':' + ('00' + date.getSeconds()).slice(-2);
    if ((DEBUG & 2) && getScrn) {
      filename = debugScrnFile + '_' +
        ('000' + (debugScrnCnt++)).slice(-3) + '.png';
      page.render(filename);
    }
    console.log('[DEBUG ' + timeStamp + '] ' + msg +
      (filename ? ' <SCREEN>: ' + filename : ''));
  }

  function pageInit() {
    // Don't use DOMContentLoaded event because script in the page is first.
    page.onLoadFinished = function(status) {
      debugLog('onLoadFinished() called. Status: ' + status + ' URL: ' + page.url);
      if (status === 'success') {
        pageInit(); // For next loading. Now, onLoadFinished was reset.
        if (jqLoad) { // Inject jQuery to new page.
          if (!page.evaluate(
              function() { return window.jQuery ? true : false; })) {
            debugLog('Inject jQuery.');
            if (!page.injectJs(JQ_PATH)) {
              console.error('Can\'t load "' + JQ_PATH + '"');
              phantom.exit(1);
            }
          } else { debugLog('jQuery already exists.'); }
        }
        if (onReadyCallback) {
          setTimeout(function() {
            onReadyCallback();
            // onLoadFinished may be called again at same page.
            onReadyCallback = null;
          }, 0);
        }
      } else if (!ignoreFail) {
        console.error('Unable to load the address! : ' + page.url);
        phantom.exit(1);
      }
    };
  }

  // pageOpen() instead of page.open().
  function pageOpen(url, onReady) {
    onReadyCallback = onReady;
    page.open(url);
  }

  // action: URI or Function that returns boolean,
  //          or [URI, jqLoad, ignoreFail] or [Function, jqLoad, ignoreFail].
  // examine: Function that returns *true* to next step or String to exit.
  function dfdAction(name, action, examine, redoInterval) {
    var dfd = $.Deferred(), timeCount = TIMEOUT,
      actionDone = false, redoCount = redoInterval;
    clearTimeout(timerId); // No Passing
    function docReady() {
      return page.evaluate(function(jqLoad) {
        return document.readyState === 'complete' && (!jqLoad || window.jQuery);
      }, jqLoad);
    }
    function actionLoop() {
      var examineRes = 0;
      debugLog(name + ' ---- actionLoop() Try: ' + (TIMEOUT - timeCount + 1), true);
      // The page may be loaded by examine().
      if (docReady() && (examineRes = examine()) === true) {
        debugLog(name + ' <EXAMINE>: OK');
        debugLog(name + ' ======== /dfdAction()');
        dfd.resolve();
      } else if (typeof examineRes === 'string') {
        debugLog(name + ' <FAILED>');
        debugLog(name + ' ======== /dfdAction()');
        dfd.reject(name + ' ' + examineRes);
      } else if ((timeCount -= countStep) <= 0) {
        debugLog(name + ' <TIMEOVER>');
        debugLog(name + ' ======== /dfdAction()');
        dfd.reject(name + ' Can\'t parse page');
      } else if (docReady() && !actionDone) {
        if (Array.isArray(action)) {
          if (typeof action[1] === 'boolean') { jqLoad = action[1]; }
          if (typeof action[2] === 'boolean') { ignoreFail = action[2]; }
          // Not array in next loop
          action = action[0];
        }
        if (typeof action === 'string') {
          debugLog(name + ' <ACTION-OPEN>: ' + action);
          actionDone = true;
          pageOpen(action, actionLoop);
        } else if (typeof action === 'function') {
          debugLog(name + ' <ACTION-FUNCTION>');
          actionDone = action();
          clearTimeout(timerId); // No Passing
          timerId = setTimeout(actionLoop, actionDone ? 1 : actionLoopInterval);
        }
      } else {
        if (actionDone && redoInterval && --redoCount <= 0) {
          debugLog(name + ' <REDO-ACTION>');
          actionDone = false;
          redoCount = redoInterval;
        }
        clearTimeout(timerId); // No Passing
        timerId = setTimeout(actionLoop, actionLoopInterval);
      }
    }

    debugLog(name + ' ======== dfdAction()');
    actionLoop();
    return dfd.promise();
  }

  // page.settings.loadImages = false;
  // If page.settings.webSecurityEnabled isn't set to false,
  // XMLHttpRequest can't do cross-domain. But, it seem to no problem?
  // Or, Office Web Apps need it?
  page.settings.webSecurityEnabled = false;
  // page.settings.localToRemoteUrlAccessEnabled = true;
  // MS may kick PhantomJS someday.
  page.settings.userAgent =
    (page.settings.userAgent + '').replace(/\s*PhantomJS(?:\/[\d\.]+)?/i, '');
  page.customHeaders = defaultReqHeaders;
  // Inject jQuery to this space and initial page.
  if (!phantom.injectJs(JQ_PATH) || !page.injectJs(JQ_PATH)) {
    console.error('Can\'t load "' + JQ_PATH + '"');
    phantom.exit(1);
  }

  pageInit();
  // "jQuery" must be used instead of "$" in the page. $ is undefined.
  // Don't use jQuery in navigation.
  // ============================ Step 00: Open SkyDrive login page
  dfdAction('Step 00',
    'https://skydrive.live.com/',
    // Current PhantomJS can't parse cookie jar and persistent login.
    // Therefore, do login every time.
    function() {
      return page.evaluate(function(selector) {
        return document.getElementById(selector.loginUser) &&
          document.getElementById(selector.loginPw) &&
          document.getElementById(selector.loginBtn) ? true : false;
      }, selector);
    }
  )
  // ============================ Step 01: Login
  .then(function() { return dfdAction('Step 01',
    function() {
      return page.evaluate(function(selector, msUser, msPw) {
        console.log('Step 01: Submit login form');
        document.getElementById(selector.loginUser).value = msUser;
        document.getElementById(selector.loginPw).value = msPw;
        // // Persistent login
        // document.getElementById(selector.loginPersistent).checked = true;
        jQuery('#' + selector.loginBtn).trigger('click');
        return true;
      }, selector, msUser, msPw);
    },
    function() {
      return page.evaluate(function(selector/*, msUser*/) {
        // if (!document.getElementById(selector.logoutBtn) ||
        //     !window.$Config) { return false; }
        // // The user which already logged in (persistent login) may be another user.
        // if (window.$Config.email === msUser) { return true; }
        return document.getElementById(selector.logoutBtn) ? true : false;
      }, selector/*, msUser*/);
    }
  ); })
  // ============================ Step 02: Open "File" page
  .then(function() { return dfdAction('Step 02',
    function() {
      return page.evaluate(function() {
        // var elmTarget = document.querySelector('.c-LeftNavBar'),
        //   navCtrl = elmTarget.control;
        // if (elmTarget && navCtrl && navCtrl.controlName === 'LeftNavBar' &&
        //     navCtrl.dataContext &&
        //     navCtrl.dataContext.item &&
        //     navCtrl.dataContext.item.urls &&
        //     typeof navCtrl.dataContext.item.urls.viewInBrowser === 'string') {
        //   document.location = navCtrl.dataContext.item.urls.viewInBrowser;
        //   console.log('Step 02: ' + navCtrl.dataContext.item.urls.viewInBrowser);
        //   return true;
        // } else { return false; }
        var jqTarget =
          jQuery('.c-LeftNavBar>.quickview:first a.quickview_header_text');
        if (jqTarget.length) {
          console.log('Step 02: Click the navigation to root directory');
          jqTarget.trigger('click');
          return true;
        } else {
          console.log('Step 02: The navigation to root directory is not found');
          return false;
        }
      });
    },
    function() {
      return page.evaluate(function(selector) {
        return document.querySelector(selector.uploadElm) ? true : false;
      }, selector);
    }
  ); })
  // ============================ Step 03: Get upload file name
  .then(function() { return dfdAction('Step 03',
    function() {
      if (page.evaluate(function()
          { return document.getElementsByClassName('c-ListView'); })) {
        debugLog('Step 03: Upload file "' + tmpFile.baseName + '" already exists');
        tmpFile.reNew();
        return true;
      } else {
        debugLog('Step 03: The file list not ready');
        return false;
      }
    },
    function() {
      // RegExp is not primitive type, but it's converted to literal code.
      // https://github.com/ariya/phantomjs/blob/269154071100332c201fc619b658b07d9fdd6cd6/src/modules/webpage.js#L375
      return page.evaluate(function(selector, re) {
        return document.getElementsByClassName('c-ListView') &&
          !jQuery(selector.fileItem).filter(function() {
            return re.test(jQuery(this).text());
          }).length ? true : false;
      }, selector, tmpFile.baseNameRe);
    }
  ); })
  // ============================ Step 04: File upload
  .then(function() { return dfdAction('Step 04',
    function() {
      if (page.evaluate(function() {
            return window.wLive &&
              window.wLive.Core &&
              window.wLive.Core.Html5FileUpload &&
              window.$Network && window.$Network.fetchXML;
          })) {
        try {
          // fs.copy returns false always?
          fs.copy(srcFile, tmpFile.path);
          // if (!fs.copy(srcFile, tmpFile.path)) {
          if (!fs.exists(tmpFile.path) || !fs.isFile(tmpFile.path)) {
            console.error('Can\'t copy file: "' + srcFile +
              '" to "' + tmpFile.path + '"');
            phantom.exit(1);
          }
        } catch(e) {
          console.error('Error(' + e.message + ') ' + e.description + ' - ' + e);
          phantom.exit(1);
        }
        debugLog('Step 04: File upload: "' + tmpFile.path + '"');
        page.evaluate(function() {
/*
The Blob is disabled at XMLHttpRequest.send in PhantomJS.
https://github.com/ariya/phantomjs/blob/269154071100332c201fc619b658b07d9fdd6cd6/src/qt/src/3rdparty/webkit/Source/WebCore/xml/XMLHttpRequest.cpp#L550
  Calling Order
  1. wLive.Core.Html5FileUpload
  2. wLive.Core.SingleFileUpload
  3. wLive.Core.Html5FileUpload.post
  4. $Network.fetchXML
  5. Some frames are loaded...
  6. <iframe>.contentWindow.fetchXML
  7. XMLHttpRequest.send (in iframe)
*/
          if (!window._fetchXML) {
            window._fetchXML = window.$Network.fetchXML;
            // Some iframes may be loaded later, and this must be repeated.
            window.$Network.fetchXML = function(url, cb, type, data, h, e, o) {
              jQuery('iframe').each(function(i, elm) {
                if (elm.contentWindow.fetchXML && !elm.contentWindow._xhrSend) {
                  elm.contentWindow._xhrSend =
                    elm.contentWindow.XMLHttpRequest.prototype.send;
                  elm.contentWindow.XMLHttpRequest.prototype.send =
                    function(data) {
                      var me = this, reader;
                      if (data instanceof Blob) { // Blob includes File.
                        console.log('Step 04: XMLHttpRequest.send: Blob(' +
                          (data.size || 'None') + ') to ArrayBuffer');
                        reader = new FileReader();
                        reader.onload = function(e) {
                          elm.contentWindow._xhrSend.call(me, e.target.result);
                        };
                        reader.readAsArrayBuffer(data);
                      } else {
                        elm.contentWindow._xhrSend.call(me, data);
                      }
                    };
                }
              });
              return window._fetchXML(url, cb, type, data, h, e, o);
            };
          }
        });
        page.uploadFile(selector.uploadElm, [tmpFile.path]);
        return true;
      } else {
        debugLog('Step 04: The file upload scripts not ready');
        return false;
      }
    },
    function() {
      var res = page.evaluate(function(selector, re) {
        var jqError = jQuery('#processMenuContainer .procr_cont.procr_error')
          .filter(function() { // children() is not safe.
            return re.test(jQuery(this).find('.procr_filename').text());
          });
        if (jqError.length) {
          return jqError.find('.procr_errorText .procr_error').text() || 'Error';
        }
        return jQuery(selector.fileItem).filter(function() {
            return re.test(jQuery(this).text());
          }).length ? true : false;
      }, selector, tmpFile.baseNameRe);
      if (res) { // It may be error message.
        uriFiles = page.url;
        fs.remove(tmpFile.path);
      }
      return res;
    }
  ); })
  // ============================ Step 05: Select document
  // If click the document direct, it's opened as edit mode.
  .then(function() { return dfdAction('Step 05',
    function() {
      debugLog('Step 05: Select document');
      page.evaluate(function(selector, re) {
        jQuery(selector.fileItem).filter(function() {
            return re.test(jQuery(this).text());
          }).trigger('contextmenu');
      }, selector, tmpFile.baseNameRe);
      return true;
    },
    function() {
      return (uriAppFrame = page.evaluate(function(selector) {
        return jQuery(selector.cmdView).filter(function() {
            return (/^open\b.+\bweb app$/i).test(jQuery(this).text());
          }).prop('href');
      }, selector)) ? true : false;
    }
  ); })
  // ============================ Step 06: Open file via Office Web Apps
  .then(function() { return dfdAction('Step 06',
    function() {
      if (uriApp) {
        if (page.evaluate(function()
            { return document.readyState === 'complete' && !window.jQuery; })) {
          debugLog('Step 06: Inject jQuery into AppFrame.');
          if (!page.injectJs(JQ_PATH)) {
            console.error('Can\'t load "' + JQ_PATH + '"');
            phantom.exit(1);
          }
        }
        if (page.evaluate(function() {
              return document.body.id === 'MainApp' && window.jQuery &&
                window.XMLHttpRequest && window.XMLHttpRequest.prototype._open;
            })) {
          jqLoad = true; // For examine
          ignoreFail = false;
          page.onInitialized = undefined;
          return true;
        }
      } else if (uriAppFrame) { // First action
        debugLog('Step 06: Start "... Web App"');
        jqLoad = false; // jQuery disturb loading.
        actionLoopInterval = 50; // Find up URL fast.
        countStep = 0.05;
        ignoreFail = true;
        pageOpen(uriAppFrame);
        uriAppFrame = '';
      } else {
        uriApp = page.evaluate(function(selector) {
          var elmTarget, elms = document.getElementsByTagName('iframe'), i;
          for (i = elms.length - 1; i >= 0; i--) {
            elms[i].contentWindow.stop();
            if (elms[i].id === selector.frameApp) {
              elmTarget = elms[i];
            }
          }
          return elmTarget ? elmTarget.src : '';
        }, selector);
        if (uriApp) {
          debugLog('Step 06: Unframe Web App');
          actionLoopInterval = 1000;
          countStep = 1;
          page.onInitialized = function() {
/*
  XMLHttpRequest removes "Content-Type" request header
  which set via XMLHttpRequest.setRequestHeader when "GET" method is specified.
  Or, it's not accepted via XMLHttpRequest.setRequestHeader? (QtWebkit/PhantomJS bug?)
  If "Content-Type" isn't set, below API returns error.
  https://excel.officeapps.live.com/x/16.0.1727.1035/_vti_bin/DynamicGridContent.json/GetRangeContent
  And, WebPage.onResourceRequested can't set headers now.
  WebPage.customHeaders is copied in
    XMLHttpRequest.send > NetworkAccessManager.createRequest.
  And, WebPage.onResourceRequested is called, after copying.
  https://github.com/ariya/phantomjs/blob/4989445e714bb97fe49ef8c00ccbce8b6cfbfbb0/src/qt/src/3rdparty/webkit/Source/WebCore/xml/XMLHttpRequest.cpp#L536
  https://github.com/ariya/phantomjs/blob/4989445e714bb97fe49ef8c00ccbce8b6cfbfbb0/src/networkaccessmanager.cpp#L230
*/
            debugLog('Step 06: Setup customized XMLHttpRequest.');
            page.onCallback = function(reqHeaders) {
              page.customHeaders = jQuery.extend(reqHeaders, defaultReqHeaders);
            };
            page.evaluate(function() {
              window.XMLHttpRequest.prototype._open = window.XMLHttpRequest.prototype.open;
              window.XMLHttpRequest.prototype.open = function(method, url, async, user, password) {
                this._method = method.toLowerCase();
                this._reqHeaders = {};
                this._open(method, url, async, user, password);
              };
              window.XMLHttpRequest.prototype._setRequestHeader =
                window.XMLHttpRequest.prototype.setRequestHeader;
              window.XMLHttpRequest.prototype.setRequestHeader = function(header, value) {
                if (this._method === 'get') {
                  this._reqHeaders[header] = value;
                } else {
                  this._setRequestHeader(header, value);
                }
              };
              window.XMLHttpRequest.prototype._send = window.XMLHttpRequest.prototype.send;
              window.XMLHttpRequest.prototype.send = function(body) {
                if (this._method === 'get') {
                  window.callPhantom(this._reqHeaders);
                }
                this._send(body);
              };
            });
          };
          pageOpen(uriApp);
        }
      }
      return false;
    },
    function() {
      return page.evaluate(function(selector) {
        var jqMsgCtrl;
        if (!window.jQuery) { return false; }
        // Warning message: the following features will be removed...
        // Any messages which have "Continue" button.
        // The class name 'ewa...' may be excel.
        if ((jqMsgCtrl = jQuery('div[role="dialog"] button[type="submit"]')).length &&
            jqMsgCtrl.text().toLowerCase() === 'continue') {
          console.log('Step 06: Click "Continue" button in dialog.');
          jqMsgCtrl.trigger('click');
          return false;
        }
        // Feedback message: Would you like to participate...?
        if ((jqMsgCtrl = jQuery('td>table a[href="Close"]')).length) {
          console.log('Step 06: Click "Close" button in dialog.');
          jqMsgCtrl.trigger('click');
          return false;
        }
        return jQuery(selector.cmdFile).length ? true : false;
      }, selector);
    }
  ); })
  // ============================ Step 07: Click "FILE", "Print"
  .then(function() { return dfdAction('Step 07',
    function() {
      page.evaluate(function() {
        if (window.nativeEvent) { return; }
        window.nativeEvent = function(jqTarget, type) {
          var offset = jqTarget.offset(),
            newEvent = document.createEvent('MouseEvents');
          newEvent.initMouseEvent(type, true, true, window, 0,
            offset.left + 100, offset.top + 100, offset.left + 10, offset.top + 10,
            false, false, false, false, 0, null);
          jqTarget.get(0).dispatchEvent(newEvent);
        };
      });
      if (!appClass) {
        appClass = page.evaluate(function(selector) {
          if (!window.clickF || window.clickF++ > 3) {
            console.log('Step 07: Click "FILE"');
            // Layout broken sometime.
            jQuery('#stripLeft,#m_excelWebRenderer_ewaCtl_stripLeft,' +
              '#PptUpperToolbar\\.LeftButtonDock').hide();
            window.nativeEvent(jQuery(selector.cmdFile), 'mousedown');
            window.clickF = 1;
          }
          return jQuery(selector.cmdPrint.word).length ? 'word' :
            jQuery(selector.cmdPrint.excel).length ? 'excel' :
            jQuery(selector.cmdPrint.ppoint).length ? 'ppoint' : false;
        }, selector);
        return false;
      }
      return page.evaluate(function(selector, appClass) {
        var jqTarget = jQuery(selector.cmdPrint[appClass]);
        if (jqTarget.length) {
          if (parseInt(jqTarget.css('opacity'), 10) === 1) {
            console.log('Step 07: Click "Print" Menu (application: ' + appClass + ')');
            window.nativeEvent(jqTarget, 'mousedown');
            window.clickF = 0; // For REDO-ACTION
            return true;
          }
        } else  if (!window.clickF || window.clickF++ > 3) {
          console.log('Step 07: Click "FILE" (application: ' + appClass + ')');
          // Layout broken sometime.
          jQuery('#stripLeft,#m_excelWebRenderer_ewaCtl_stripLeft,' +
            '#PptUpperToolbar\\.LeftButtonDock').hide();
          window.nativeEvent(jQuery(selector.cmdFile), 'mousedown');
          window.clickF = 1;
        }
        return false;
      }, selector, appClass);
    },
    function() {
      return appClass && page.evaluate(function(selector, appClass) {
        var jqTarget = jQuery(selector.cmdPrintSub[appClass]);
        if (!jqTarget.length) { return false; }
        if (jqTarget.attr('aria-disabled') === 'true') {
          // retry
          jqTarget = jQuery('#jbtnBackArrow-Menu32');
          window.nativeEvent(jqTarget, 'mousedown');
          window.setTimeout(function() { window.nativeEvent(jqTarget, 'mouseup'); }, 0);
          return false;
        }
        return parseInt(jqTarget.css('opacity'), 10) === 1 &&
          jQuery('.cui-jewelsubmenu').length === 1 ? true : false;
      }, selector, appClass);
    },
    3 // Menu showing may be reseted.
  ); })
  // ============================ Step 08: Click "Print" Sub Menu
  .then(function() { return dfdAction('Step 08',
    function() {
      debugLog('Step 08: Click "Print" Sub Menu (application: ' + appClass + ')');
      page.evaluate(function(selector, appClass) {
        (function() {
          var jqTarget = jQuery(selector.cmdPrintSub[appClass]);
          window.nativeEvent(jqTarget, 'mousedown');
          window.setTimeout(function() { window.nativeEvent(jqTarget, 'mouseup'); }, 0);
        })();
      }, selector, appClass);
      return true;
    },
    function() {
      if (appClass === 'excel') {
        // ==================== excel
        return page.evaluate(function(selector) {
          var jqDialog, jqBtn;
          return (jqDialog = jQuery(selector.cmdExcel.dialog)).length &&
            jqDialog.find(selector.cmdExcel.optEntire).length &&
            (jqBtn = jqDialog.find(selector.cmdExcel.btnPrint)).length &&
            !jqBtn.prop('disabled') ? true : false;
        }, selector);
      } else {
        // ==================== others
        return page.evaluate(function(selector) {
          return jQuery(selector.linkPdf).length ? true : false;
        }, selector);
      }
    }
  ); })
  // ============================ Step 09: Get PDF
  .then(function() { return dfdAction('Step 09',
    [function() {
      debugLog('Step 09: Get PDF (application: ' + appClass + ')');
      var uri;
      if (appClass === 'excel') {
        // ==================== excel
        page.onPageCreated = function(newPage) {
          debugLog('onPageCreated Step 09: Save PDF (application: ' + appClass + ')');
          if (DEBUG) {
            newPage.onError = page.onError;
            if (DEBUG & 4) {
              newPage.onResourceReceived = page.onResourceReceived;
            }
          }
          newPage.onResourceRequested = page.onResourceRequested;
          newPage.onResourceError = page.onResourceError;
          newPage.onLoadFinished = function(status) {
            var filename;
            if (DEBUG & 2) {
              filename = debugScrnFile + '_' +
                ('000' + (debugScrnCnt++)).slice(-3) + '.png';
              newPage.render(filename);
              debugLog('(newPage) onLoadFinished Step 09: Status: ' + status +
                ' <SCREEN>: ' + filename);
            }
            if (status === 'success') {
              if (newPage.evaluate(function(selector) {
                    var elms = document.querySelectorAll(selector.cmdExcel.ctrlArea), i;
                    if (!elms.length) { return false; }
                    for (i = elms.length - 1; i >= 0; i--) {
                      elms[i].style.display = 'none';
                    }
                    return true;
                  }, selector)) {
                newPage.paperSize =
                  { format: 'A4', orientation: 'portrait', margin: '1cm' };
                window.setTimeout(function () {
                  try {
                    newPage.render(destFile);
                    if (!fs.exists(destFile) || !fs.isFile(destFile)) {
                      console.error('Can\'t save file: "' + destFile + '"');
                      phantom.exit(1);
                    }
                  } catch(e) {
                    console.error('Error(' + e.message + ') ' +
                      e.description + ' - ' + e);
                    phantom.exit(1);
                  }
                  newPage.close();
                  pdfSaved = true;
                }, 200);
              }/* else {
                console.error('Can\'t get view');
                phantom.exit(1);
              }*/
            }/* else {
              console.error('Unable to load the address! : ' + newPage.url);
              phantom.exit(1);
            }*/
          };
        };

        page.evaluate(function(selector) {
          var jqDialog = jQuery(selector.cmdExcel.dialog);
          jqDialog.find(selector.cmdExcel.optEntire).prop('checked', true);
          jqDialog.find(selector.cmdExcel.btnPrint).click();
        }, selector);
        return true;
      } else {
        // ==================== others
        uri = page.evaluate(function(selector)
          { return jQuery(selector.linkPdf).prop('href'); }, selector);
        if (!uri) { return false; }

        page.onCallback = function(data) {
          debugLog('onCallback Step 09: Save PDF (application: ' + appClass + ')');
          try {
            // fs.write returns false always?
            fs.write(destFile, atob(data.base64), 'wb');
            // if (!fs.write(destFile, atob(data.base64), 'wb')) {
            if (!fs.exists(destFile) || !fs.isFile(destFile)) {
              console.error('Can\'t save file: "' + destFile + '"');
              phantom.exit(1);
            }
          } catch(e) {
            console.error('Error(' + e.message + ') ' + e.description + ' - ' + e);
            phantom.exit(1);
          }
          pdfSaved = true;
        };

        page.evaluate(function(uri) {
          // Restore
          window.XMLHttpRequest.prototype.open = window.XMLHttpRequest.prototype._open;
          window.XMLHttpRequest.prototype.setRequestHeader =
            window.XMLHttpRequest.prototype._setRequestHeader;
          window.XMLHttpRequest.prototype.send = window.XMLHttpRequest.prototype._send;

// ---------------- base64ArrayBuffer
// https://gist.github.com/958841
function base64ArrayBuffer(arrayBuffer) {
  var base64    = '';
  var encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

  var bytes         = new Uint8Array(arrayBuffer);
  var byteLength    = bytes.byteLength;
  var byteRemainder = byteLength % 3;
  var mainLength    = byteLength - byteRemainder;

  var a, b, c, d;
  var chunk;

  // Main loop deals with bytes in chunks of 3
  for (var i = 0; i < mainLength; i = i + 3) {
    // Combine the three bytes into a single integer
    chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];

    // Use bitmasks to extract 6-bit segments from the triplet
    a = (chunk & 16515072) >> 18; // 16515072 = (2^6 - 1) << 18
    b = (chunk & 258048)   >> 12; // 258048   = (2^6 - 1) << 12
    c = (chunk & 4032)     >>  6; // 4032     = (2^6 - 1) << 6
    d = chunk & 63;               // 63       = 2^6 - 1

    // Convert the raw binary segments to the appropriate ASCII encoding
    base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
  }

  // Deal with the remaining bytes and padding
  if (byteRemainder === 1) {
    chunk = bytes[mainLength];

    a = (chunk & 252) >> 2; // 252 = (2^6 - 1) << 2

    // Set the 4 least significant bits to zero
    b = (chunk & 3)   << 4; // 3   = 2^2 - 1

    base64 += encodings[a] + encodings[b] + '==';
  } else if (byteRemainder === 2) {
    chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];

    a = (chunk & 64512) >> 10; // 64512 = (2^6 - 1) << 10
    b = (chunk & 1008)  >>  4; // 1008  = (2^6 - 1) << 4

    // Set the 2 least significant bits to zero
    c = (chunk & 15)    <<  2; // 15    = 2^4 - 1

    base64 += encodings[a] + encodings[b] + encodings[c] + '=';
  }

  return base64;
}
// ---------------- /base64ArrayBuffer

          var xhr = new XMLHttpRequest();
          xhr.open('GET', uri, true);
          xhr.responseType = 'arraybuffer';
          xhr.onreadystatechange = function() {
            if (xhr.readyState === 4 && xhr.status === 200){
              window.callPhantom({base64: base64ArrayBuffer(xhr.response)});
            }
          };
          xhr.send();
        }, uri);
        return true;
      }
    }, false],
    function() {
      return pdfSaved;
    }
  ); })
  // ============================ Step 10: Back to file-list
  .then(function() { return dfdAction('Step 10',
    [uriFiles, true, true], // jqLoad to on
    function() {
      return page.evaluate(function(selector, re) {
        return window.jQuery && jQuery('#m_wh').css('display') === 'block' &&
          jQuery(selector.fileItem).filter(function() {
            return re.test(jQuery(this).text());
          }).length ? true : false;
      }, selector, tmpFile.baseNameRe);
    }
  ); })
  // ============================ Step 11: Select document
  .then(function() { return dfdAction('Step 11',
    [function() {
      debugLog('Step 11: Select document');
      page.evaluate(function(selector, re) {
        jQuery(selector.fileItem).filter(function() {
            return re.test(jQuery(this).text());
          }).trigger('contextmenu');
      }, selector, tmpFile.baseNameRe);
      return true;
    }, true, false],
    function() {
      return page.evaluate(function(selector) {
        return jQuery(selector.cmdRemove).filter(function() {
            return jQuery(this).text().toLowerCase() === 'delete';
          }).length ? true : false;
      }, selector);
    }
  ); })
  // ============================ Step 12: Remove file
  .then(function() { return dfdAction('Step 12',
    function() {
      debugLog('Step 12: Click "Delete"');
      page.evaluate(function(selector) {
        jQuery(selector.cmdRemove).filter(function() {
            return jQuery(this).text().toLowerCase() === 'delete';
          }).trigger('click');
      }, selector);
      return true;
    },
    function() {
      return page.evaluate(function(selector, re) {
        return document.getElementsByClassName('c-ListView') &&
          !jQuery(selector.fileItem).filter(function() {
            return re.test(jQuery(this).text());
          }).length ? true : false;
      }, selector, tmpFile.baseNameRe);
    }
  ); })
  // ============================ DONE
  .done(function() {
    console.log('Done.\napplication: ' + appClass + '\nfile: "' + destFile + '"');
    phantom.exit();
  })
  .fail(function(msg) {
    console.error('ERROR: ' + msg);
    phantom.exit(1);
  });

})();
