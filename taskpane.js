/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "cedaf607fc3782e05ca5.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/^blob:/, "").replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = document.baseURI || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be in strict mode.
!function() {
"use strict";
var __webpack_exports__ = {};
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
// Module
var code = "<!DOCTYPE html>\n<html>\n\n<head>\n    <meta charset=\"UTF-8\" />\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n    <title>웍스AI 엑셀 도우미</title>\n\n    <!-- Office JavaScript API -->\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js\"><" + "/script>\n\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\n    <link rel=\"stylesheet\" href=\"https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css\" />\n    <link rel=\"stylesheet\" href=\"https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css\" />\n\n    <!-- Template styles -->\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\n</head>\n\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\n    <main id=\"app-body\" class=\"ms-welcome__main\">\n        <div class=\"container\">\n            <div class=\"header-section\">\n                <h2 class=\"ms-font-xl\">웍스AI 엑셀 도우미</h2>\n                <p class=\"help-text\">자연어로 Excel 작업을 요청하세요. (대용량 번역 지원)</p>\n            </div>\n\n            <!-- Voice Input Section -->\n            <div class=\"voice-section\">\n                <button id=\"voiceButton\" class=\"voice-button\">\n                    <svg width=\"24\" height=\"24\" viewBox=\"0 0 24 24\" xmlns=\"http://www.w3.org/2000/svg\">\n                        <path d=\"M12 14c1.66 0 3-1.34 3-3V5c0-1.66-1.34-3-3-3S9 3.34 9 5v6c0 1.66 1.34 3 3 3z\"/>\n                        <path d=\"M17 11c0 2.76-2.24 5-5 5s-5-2.24-5-5H5c0 3.53 2.61 6.43 6 6.92V21h2v-3.08c3.39-.49 6-3.39 6-6.92h-2z\"/>\n                    </svg>\n                </button>\n                <div class=\"voice-status\" id=\"voiceStatus\">음성 인식 준비</div>\n            </div>\n\n            <div class=\"input-container\">\n                <textarea \n                    id=\"commandInput\" \n                    placeholder=\"예: A1부터 A10까지 합계를 구해줘\"\n                    rows=\"3\"\n                ></textarea>\n                \n                <div class=\"button-container\">\n                    <button id=\"executeButton\" class=\"ms-Button ms-Button--primary\">\n                        <span class=\"ms-Button-label\">실행</span>\n                    </button>\n                    <button id=\"clearButton\" class=\"ms-Button\">\n                        <span class=\"ms-Button-label\">지우기</span>\n                    </button>\n                </div>\n            </div>\n\n            <div id=\"status\" class=\"status-message\"></div>\n\n            <div class=\"examples\">\n                <h3>예제 명령어</h3>\n                <div class=\"example-list\">\n                    <div class=\"example-item\" data-command=\"A1:B5 셀을 병합해줘\">\n                        <strong>셀 병합:</strong> A1:B5 셀을 병합해줘\n                    </div>\n                    <div class=\"example-item\" data-command=\"A열의 합계를 구해줘\">\n                        <strong>합계 계산:</strong> A열의 합계를 구해줘\n                    </div>\n                    <div class=\"example-item\" data-command=\"Name 열의 합계를 구해줘\">\n                        <strong>레이블로 합계:</strong> Name 열의 합계를 구해줘\n                    </div>\n                    <div class=\"example-item\" data-command=\"선택한 셀을 굵게 만들고 파란색으로 바꿔줘\">\n                        <strong>서식 지정:</strong> 선택한 셀을 굵게 만들고 파란색으로 바꿔줘\n                    </div>\n                    <div class=\"example-item\" data-command=\"B열 기준으로 내림차순 정렬해줘\">\n                        <strong>정렬:</strong> B열 기준으로 내림차순 정렬해줘\n                    </div>\n                    <div class=\"example-item\" data-command=\"값이 100보다 큰 셀은 녹색 배경으로 표시해줘\">\n                        <strong>조건부 서식:</strong> 값이 100보다 큰 셀은 녹색 배경으로 표시해줘\n                    </div>\n                    <div class=\"example-item\" data-command=\"A1:B10 데이터로 막대 차트를 만들어줘\">\n                        <strong>차트 생성:</strong> A1:B10 데이터로 막대 차트를 만들어줘\n                    </div>\n                    <div class=\"example-item\" data-command=\"D열을 중국어로 번역해서 다음 열에 추가해줘\">\n                        <strong>번역:</strong> D열을 중국어로 번역해서 다음 열에 추가해줘\n                    </div>\n                    <div class=\"example-item\" data-command=\"A열의 10000개 행을 영어로 번역해줘\">\n                        <strong>대용량 번역:</strong> A열의 10000개 행을 영어로 번역해줘\n                    </div>\n                    <div class=\"example-item\" data-command=\"D2:D170 사이의 빈 행을 제거해줘\">\n                        <strong>빈 행 제거:</strong> D2:D170 사이의 빈 행을 제거해줘\n                    </div>\n                </div>\n            </div>\n\n            <div class=\"settings-section\">\n                <button id=\"settingsButton\" class=\"settings-link\">설정</button>\n            </div>\n        </div>\n    </main>\n</body>\n\n</html>";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Excel */

// Voice recognition variables
var recognition = null;
var isListening = false;

// Backend API URL - Vercel 배포 후 실제 URL로 변경해주세요
// 디버그용 URL 사용 중 (CORS 문제 해결 후 원래 URL로 변경 필요)
var API_PROXY_URL = "http://localhost:3000/api/openai-proxy" || 0;
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    // Test backend connection
    testBackendConnection();

    // Assign event handlers
    document.getElementById("executeButton").onclick = executeCommand;
    document.getElementById("clearButton").onclick = clearInput;
    document.getElementById("voiceButton").onclick = toggleVoiceRecognition;
    document.getElementById("settingsButton").onclick = showSettings;

    // Handle example clicks
    var exampleItems = document.querySelectorAll('.example-item');
    exampleItems.forEach(function (item) {
      item.onclick = function () {
        setCommand(this.getAttribute('data-command'));
      };
    });

    // Handle Enter key
    document.getElementById('commandInput').addEventListener('keydown', function (event) {
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        executeCommand();
      }
    });

    // Initialize speech recognition
    initializeSpeechRecognition();
  }
});

// Initialize speech recognition
function initializeSpeechRecognition() {
  var SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    showStatus('음성 인식이 지원되지 않는 브라우저입니다.', 'error');
    document.getElementById('voiceButton').disabled = true;
    document.getElementById('voiceStatus').textContent = '음성 인식 미지원';
    return null;
  }
  recognition = new SpeechRecognition();
  recognition.lang = 'ko-KR';
  recognition.continuous = false;
  recognition.interimResults = true;
  recognition.maxAlternatives = 1;
  recognition.onstart = function () {
    isListening = true;
    document.getElementById('voiceButton').classList.add('listening');
    document.getElementById('voiceStatus').textContent = '듣고 있습니다... 말씀해주세요';
    document.getElementById('voiceStatus').classList.add('listening');
  };
  recognition.onresult = function (event) {
    var transcript = event.results[0][0].transcript;
    document.getElementById('commandInput').value = transcript;
    if (event.results[0].isFinal) {
      document.getElementById('voiceStatus').textContent = '음성 인식 완료';
    } else {
      document.getElementById('voiceStatus').textContent = '인식중: ' + transcript;
    }
  };
  recognition.onerror = function (event) {
    isListening = false;
    document.getElementById('voiceButton').classList.remove('listening');
    document.getElementById('voiceStatus').classList.remove('listening');
    var errorMessage = '음성 인식 오류';
    switch (event.error) {
      case 'no-speech':
        errorMessage = '음성이 감지되지 않았습니다.';
        break;
      case 'audio-capture':
        errorMessage = '마이크를 찾을 수 없습니다.';
        break;
      case 'not-allowed':
        errorMessage = '마이크 권한이 거부되었습니다.';
        break;
      case 'network':
        errorMessage = '네트워크 오류가 발생했습니다.';
        break;
    }
    document.getElementById('voiceStatus').textContent = errorMessage;
    showStatus(errorMessage, 'error');
  };
  recognition.onend = function () {
    isListening = false;
    document.getElementById('voiceButton').classList.remove('listening');
    document.getElementById('voiceStatus').classList.remove('listening');
    var command = document.getElementById('commandInput').value.trim();
    if (command) {
      document.getElementById('voiceStatus').textContent = '음성 인식 완료. 실행 버튼을 눌러주세요.';
      document.getElementById('executeButton').focus();
    } else {
      document.getElementById('voiceStatus').textContent = '음성 인식 준비';
    }
  };
  return recognition;
}

// Toggle voice recognition
function toggleVoiceRecognition() {
  if (!recognition) {
    recognition = initializeSpeechRecognition();
    if (!recognition) return;
  }
  if (isListening) {
    recognition.stop();
  } else {
    recognition.start();
  }
}

// Execute command
// Add flag to prevent duplicate execution
var isExecuting = false;
function executeCommand() {
  return _executeCommand.apply(this, arguments);
} // Call OpenAI API through proxy
function _executeCommand() {
  _executeCommand = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2() {
    var command, button, originalText, _t;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.p = _context2.n) {
        case 0:
          if (!isExecuting) {
            _context2.n = 1;
            break;
          }
          console.log('Command already executing, ignoring duplicate call');
          return _context2.a(2);
        case 1:
          command = document.getElementById('commandInput').value.trim();
          if (command) {
            _context2.n = 2;
            break;
          }
          showStatus('명령어를 입력해주세요.', 'error');
          return _context2.a(2);
        case 2:
          isExecuting = true;
          button = document.getElementById('executeButton');
          originalText = '<span class="ms-Button-label">실행</span>'; // Show loading state
          button.disabled = true;
          button.innerHTML = '<span class="loading"></span><span>처리중...</span>';
          showStatus('명령을 처리하고 있습니다...', 'info');
          _context2.p = 3;
          _context2.n = 4;
          return Excel.run(/*#__PURE__*/function () {
            var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(context) {
              var worksheet, range, usedRange, sheetContext, headerRow, i, interpretation, result;
              return _regenerator().w(function (_context) {
                while (1) switch (_context.n) {
                  case 0:
                    console.log('Starting Excel.run for command execution');
                    // Get current worksheet context
                    worksheet = context.workbook.worksheets.getActiveWorksheet();
                    range = context.workbook.getSelectedRange(); // Load necessary properties
                    worksheet.load("name");
                    range.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);

                    // Get sheet data for context
                    usedRange = worksheet.getUsedRange();
                    usedRange.load(["rowCount", "columnCount", "values"]);
                    _context.n = 1;
                    return context.sync();
                  case 1:
                    // Build sheet context
                    sheetContext = {
                      sheetName: worksheet.name,
                      activeRange: {
                        address: range.address,
                        row: range.rowIndex + 1,
                        column: range.columnIndex + 1,
                        numRows: range.rowCount,
                        numColumns: range.columnCount
                      },
                      lastRow: usedRange ? usedRange.rowCount : 0,
                      lastColumn: usedRange ? usedRange.columnCount : 0,
                      headers: [],
                      dataRange: usedRange ? usedRange.values : [],
                      isLargeSheet: usedRange && usedRange.rowCount > 1000
                    }; // Extract headers
                    if (usedRange && usedRange.rowCount > 0) {
                      headerRow = usedRange.values[0];
                      for (i = 0; i < headerRow.length; i++) {
                        sheetContext.headers.push({
                          column: i + 1,
                          columnLetter: getColumnLetter(i),
                          label: headerRow[i] ? headerRow[i].toString() : ''
                        });
                      }
                    }

                    // Call OpenAI API to interpret the command
                    _context.n = 2;
                    return callOpenAI(command, sheetContext);
                  case 2:
                    interpretation = _context.v;
                    if (interpretation.success) {
                      _context.n = 3;
                      break;
                    }
                    throw new Error(interpretation.error);
                  case 3:
                    // Execute the interpreted command
                    console.log('Executing interpreted command...');
                    _context.n = 4;
                    return executeInterpretedCommand(context, interpretation.data);
                  case 4:
                    result = _context.v;
                    console.log('Command execution result:', result);

                    // Reset button and show success
                    button.disabled = false;
                    button.innerHTML = originalText;
                    if (result.success) {
                      console.log('Operation completed successfully:', result);
                      showStatus(result.message || '명령이 성공적으로 실행되었습니다.', 'success');
                      setTimeout(function () {
                        document.getElementById('commandInput').value = '';
                      }, 1000);
                    } else {
                      console.error('Operation failed:', result);
                      showStatus(result.error || '명령 실행에 실패했습니다.', 'error');
                    }
                    console.log('Excel.run completing...');
                  case 5:
                    return _context.a(2);
                }
              }, _callee);
            }));
            return function (_x38) {
              return _ref.apply(this, arguments);
            };
          }());
        case 4:
          console.log('Excel.run completed');
          _context2.n = 6;
          break;
        case 5:
          _context2.p = 5;
          _t = _context2.v;
          console.error('Error in executeCommand:', _t);
          console.error('Error stack:', _t.stack);
          button.disabled = false;
          button.innerHTML = originalText;
          showStatus('오류가 발생했습니다: ' + _t.message, 'error');
        case 6:
          _context2.p = 6;
          // Reset execution flag
          isExecuting = false;
          return _context2.f(6);
        case 7:
          return _context2.a(2);
      }
    }, _callee2, null, [[3, 5, 6, 7]]);
  }));
  return _executeCommand.apply(this, arguments);
}
function callOpenAI(_x, _x2) {
  return _callOpenAI.apply(this, arguments);
} // Original OpenAI API call (no longer used)
function _callOpenAI() {
  _callOpenAI = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3(command, sheetContext) {
    var response, errorData, _t2;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.p = _context3.n) {
        case 0:
          _context3.p = 0;
          _context3.n = 1;
          return fetch(API_PROXY_URL, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              command: command,
              sheetContext: sheetContext
            })
          });
        case 1:
          response = _context3.v;
          if (response.ok) {
            _context3.n = 3;
            break;
          }
          _context3.n = 2;
          return response.json();
        case 2:
          errorData = _context3.v;
          return _context3.a(2, {
            success: false,
            error: errorData.error || "\uC11C\uBC84 \uC624\uB958 (".concat(response.status, ")")
          });
        case 3:
          _context3.n = 4;
          return response.json();
        case 4:
          return _context3.a(2, _context3.v);
        case 5:
          _context3.p = 5;
          _t2 = _context3.v;
          console.error('Proxy API Error:', _t2);
          return _context3.a(2, {
            success: false,
            error: "API \uC694\uCCAD \uC624\uB958: ".concat(_t2.toString())
          });
      }
    }, _callee3, null, [[0, 5]]);
  }));
  return _callOpenAI.apply(this, arguments);
}
function callOpenAIDirectly(_x3, _x4) {
  return _callOpenAIDirectly.apply(this, arguments);
} // Execute the interpreted command
function _callOpenAIDirectly() {
  _callOpenAIDirectly = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4(command, sheetContext) {
    var systemPrompt, url, payload, response, _errorData$error, errorData, result, content, parsedCommand, _t3, _t4;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.p = _context4.n) {
        case 0:
          systemPrompt = "You are an Excel assistant that interprets natural language commands and returns JSON instructions for Excel operations.\n  \nAvailable operations:\n1. merge: Merge cells\n2. sum: Sum values in a range or column\n3. average: Calculate average\n4. count: Count cells (can count all, numbers only, or based on conditions)\n5. format: Format cells (bold, italic, font color, background color, etc.)\n6. sort: Sort data\n7. filter: Filter data\n8. insert: Insert rows/columns\n9. delete: Delete rows/columns\n10. formula: Add custom formula\n11. chart: Create chart\n12. conditional_format: Add conditional formatting\n13. translate: Translate cell contents to another language\n14. compress: Remove empty rows in a specific column range\n15. retry_translation: Retry translation for failed items marked as [\uBC88\uC5ED \uC2E4\uD328]\n\nFor sum operation:\n- If user mentions a column by header name (e.g., \"totalToken \uC5F4\uC758 \uD569\", \"totalToken \uD569\uC0B0\"), return: { \"sumType\": \"column\", \"columnName\": \"totalToken\" }\n- The system will automatically find the column, determine the data range, and place the sum in the first empty cell below the data\n- For specific range sum, use: { \"sourceRange\": \"A2:A10\" }\n- For adding sum below selection, use: { \"addNewRow\": true }\n\nCurrent sheet context:\n- Active range: ".concat(sheetContext.activeRange.address, "\n- Sheet dimensions: ").concat(sheetContext.lastRow, " rows x ").concat(sheetContext.lastColumn, " columns\n- Headers: ").concat(sheetContext.headers.map(function (h) {
            return "Column ".concat(h.columnLetter, ": \"").concat(h.label, "\"");
          }).join(', '), "\n\nReturn JSON in this format:\n{\n  \"operation\": \"operation_name\",\n  \"parameters\": {\n    // operation-specific parameters\n  }\n}");
          url = 'https://api.openai.com/v1/chat/completions';
          payload = {
            model: 'gpt-4.1-2025-04-14',
            messages: [{
              role: 'system',
              content: systemPrompt
            }, {
              role: 'user',
              content: "Current Excel state:\nHeaders: ".concat(sheetContext.headers.map(function (h) {
                return "Column ".concat(h.columnLetter, ": \"").concat(h.label, "\"");
              }).join(', '), "\nActive sheet: ").concat(sheetContext.sheetName, "\n\nUser command: ").concat(command)
            }],
            temperature: 0.3,
            max_tokens: 500
          };
          _context4.p = 1;
          _context4.n = 2;
          return fetch(url, {
            method: 'POST',
            headers: {
              'Authorization': "Bearer ".concat(OPENAI_API_KEY),
              'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
          });
        case 2:
          response = _context4.v;
          if (response.ok) {
            _context4.n = 6;
            break;
          }
          _context4.n = 3;
          return response.json();
        case 3:
          errorData = _context4.v;
          if (!(response.status === 429)) {
            _context4.n = 4;
            break;
          }
          return _context4.a(2, {
            success: false,
            error: 'API 요청 한도를 초과했습니다. 잠시 후 다시 시도해주세요.'
          });
        case 4:
          if (!(response.status === 401)) {
            _context4.n = 5;
            break;
          }
          return _context4.a(2, {
            success: false,
            error: 'API 키가 유효하지 않습니다. API 키를 확인해주세요.'
          });
        case 5:
          return _context4.a(2, {
            success: false,
            error: "API \uC624\uB958 (".concat(response.status, "): ").concat(((_errorData$error = errorData.error) === null || _errorData$error === void 0 ? void 0 : _errorData$error.message) || '알 수 없는 오류')
          });
        case 6:
          _context4.n = 7;
          return response.json();
        case 7:
          result = _context4.v;
          if (!(result.choices && result.choices[0])) {
            _context4.n = 10;
            break;
          }
          content = result.choices[0].message.content;
          _context4.p = 8;
          parsedCommand = JSON.parse(content);
          return _context4.a(2, {
            success: true,
            data: parsedCommand
          });
        case 9:
          _context4.p = 9;
          _t3 = _context4.v;
          console.error('Failed to parse AI response:', content);
          return _context4.a(2, {
            success: false,
            error: 'AI 응답을 해석할 수 없습니다. 다시 시도해주세요.'
          });
        case 10:
          return _context4.a(2, {
            success: false,
            error: 'OpenAI API 응답을 파싱할 수 없습니다.'
          });
        case 11:
          _context4.n = 13;
          break;
        case 12:
          _context4.p = 12;
          _t4 = _context4.v;
          console.error('OpenAI API Error:', _t4);
          return _context4.a(2, {
            success: false,
            error: "OpenAI API \uC624\uB958: ".concat(_t4.toString())
          });
        case 13:
          return _context4.a(2);
      }
    }, _callee4, null, [[8, 9], [1, 12]]);
  }));
  return _callOpenAIDirectly.apply(this, arguments);
}
function executeInterpretedCommand(_x5, _x6) {
  return _executeInterpretedCommand.apply(this, arguments);
} // Merge cells
function _executeInterpretedCommand() {
  _executeInterpretedCommand = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5(context, commandData) {
    var operation, params, _t5, _t6;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.p = _context5.n) {
        case 0:
          operation = commandData.operation;
          params = commandData.parameters || {};
          console.log("[".concat(new Date().toISOString(), "] Executing operation: ").concat(operation, " with params:"), params);
          _context5.p = 1;
          _t5 = operation;
          _context5.n = _t5 === 'merge' ? 2 : _t5 === 'sum' ? 4 : _t5 === 'average' ? 6 : _t5 === 'count' ? 8 : _t5 === 'format' ? 10 : _t5 === 'sort' ? 12 : _t5 === 'filter' ? 14 : _t5 === 'insert' ? 16 : _t5 === 'delete' ? 18 : _t5 === 'formula' ? 20 : _t5 === 'chart' ? 22 : _t5 === 'conditional_format' ? 24 : _t5 === 'translate' ? 26 : _t5 === 'compress' ? 28 : _t5 === 'retry_translation' ? 30 : 32;
          break;
        case 2:
          _context5.n = 3;
          return executeMerge(context, params);
        case 3:
          return _context5.a(2, _context5.v);
        case 4:
          _context5.n = 5;
          return executeSum(context, params);
        case 5:
          return _context5.a(2, _context5.v);
        case 6:
          _context5.n = 7;
          return executeAverage(context, params);
        case 7:
          return _context5.a(2, _context5.v);
        case 8:
          _context5.n = 9;
          return executeCount(context, params);
        case 9:
          return _context5.a(2, _context5.v);
        case 10:
          _context5.n = 11;
          return executeFormat(context, params);
        case 11:
          return _context5.a(2, _context5.v);
        case 12:
          _context5.n = 13;
          return executeSort(context, params);
        case 13:
          return _context5.a(2, _context5.v);
        case 14:
          _context5.n = 15;
          return executeFilter(context, params);
        case 15:
          return _context5.a(2, _context5.v);
        case 16:
          _context5.n = 17;
          return executeInsert(context, params);
        case 17:
          return _context5.a(2, _context5.v);
        case 18:
          _context5.n = 19;
          return executeDelete(context, params);
        case 19:
          return _context5.a(2, _context5.v);
        case 20:
          _context5.n = 21;
          return executeFormula(context, params);
        case 21:
          return _context5.a(2, _context5.v);
        case 22:
          _context5.n = 23;
          return executeChart(context, params);
        case 23:
          return _context5.a(2, _context5.v);
        case 24:
          _context5.n = 25;
          return executeConditionalFormat(context, params);
        case 25:
          return _context5.a(2, _context5.v);
        case 26:
          _context5.n = 27;
          return executeTranslate(context, params);
        case 27:
          return _context5.a(2, _context5.v);
        case 28:
          _context5.n = 29;
          return executeCompress(context, params);
        case 29:
          return _context5.a(2, _context5.v);
        case 30:
          _context5.n = 31;
          return executeRetryTranslation(context, params);
        case 31:
          return _context5.a(2, _context5.v);
        case 32:
          return _context5.a(2, {
            success: false,
            error: "\uC54C \uC218 \uC5C6\uB294 \uC791\uC5C5: ".concat(operation)
          });
        case 33:
          _context5.n = 35;
          break;
        case 34:
          _context5.p = 34;
          _t6 = _context5.v;
          console.error('Error in executeInterpretedCommand:', _t6);
          return _context5.a(2, {
            success: false,
            error: "\uC791\uC5C5 \uC2E4\uD589 \uC911 \uC624\uB958: ".concat(_t6.message || _t6.toString())
          });
        case 35:
          return _context5.a(2);
      }
    }, _callee5, null, [[1, 34]]);
  }));
  return _executeInterpretedCommand.apply(this, arguments);
}
function executeMerge(_x7, _x8) {
  return _executeMerge.apply(this, arguments);
} // Sum values
function _executeMerge() {
  _executeMerge = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6(context, params) {
    var worksheet, range, rangeAddress;
    return _regenerator().w(function (_context6) {
      while (1) switch (_context6.n) {
        case 0:
          console.log('executeMerge started with params:', params);
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          range = params.range ? worksheet.getRange(params.range) : context.workbook.getSelectedRange();
          console.log('Loading range address...');
          // Load address property before using it
          range.load('address');
          _context6.n = 1;
          return context.sync();
        case 1:
          console.log('Range address loaded:', range.address);
          rangeAddress = range.address;
          console.log('Merging range...');
          range.merge();
          _context6.n = 2;
          return context.sync();
        case 2:
          console.log('Merge completed successfully');
          return _context6.a(2, {
            success: true,
            message: "".concat(rangeAddress, " \uBC94\uC704\uAC00 \uBCD1\uD569\uB418\uC5C8\uC2B5\uB2C8\uB2E4.")
          });
      }
    }, _callee6);
  }));
  return _executeMerge.apply(this, arguments);
}
function executeSum(_x9, _x0) {
  return _executeSum.apply(this, arguments);
} // Calculate average
function _executeSum() {
  _executeSum = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee7(context, params) {
    var worksheet, usedRange, headers, columnIndex, columnLetter, i, lastDataRow, row, rangeAddress, sumCell, sourceRange, column, lastRow, newCell, targetCell, _column, _lastRow, _newCell;
    return _regenerator().w(function (_context7) {
      while (1) switch (_context7.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet(); // If column name is provided, find the column and create range
          if (!(params.columnName || params.sumType === 'column')) {
            _context7.n = 9;
            break;
          }
          console.log('Sum by column name:', params.columnName);

          // Get sheet context to find column by name
          usedRange = worksheet.getUsedRange();
          usedRange.load(["values", "rowCount", "columnCount"]);
          _context7.n = 1;
          return context.sync();
        case 1:
          if (!(!usedRange || usedRange.rowCount === 0)) {
            _context7.n = 2;
            break;
          }
          return _context7.a(2, {
            success: false,
            error: '시트에 데이터가 없습니다.'
          });
        case 2:
          // Find column by header name
          headers = usedRange.values[0];
          columnIndex = -1;
          columnLetter = '';
          i = 0;
        case 3:
          if (!(i < headers.length)) {
            _context7.n = 5;
            break;
          }
          if (!(headers[i] && headers[i].toString().toLowerCase() === (params.columnName || '').toLowerCase())) {
            _context7.n = 4;
            break;
          }
          columnIndex = i;
          columnLetter = getColumnLetter(i);
          return _context7.a(3, 5);
        case 4:
          i++;
          _context7.n = 3;
          break;
        case 5:
          if (!(columnIndex === -1)) {
            _context7.n = 6;
            break;
          }
          return _context7.a(2, {
            success: false,
            error: "\"".concat(params.columnName, "\" \uC5F4\uC744 \uCC3E\uC744 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4.")
          });
        case 6:
          // Find last row with data in this column
          lastDataRow = 1; // Start from row 2 (after header)
          for (row = 1; row < usedRange.rowCount; row++) {
            if (usedRange.values[row][columnIndex] !== null && usedRange.values[row][columnIndex] !== undefined && usedRange.values[row][columnIndex] !== '') {
              lastDataRow = row + 1; // +1 because Excel rows are 1-indexed
            }
          }

          // Create range from row 2 to last data row
          rangeAddress = "".concat(columnLetter, "2:").concat(columnLetter).concat(lastDataRow);
          sumCell = worksheet.getCell(lastDataRow, columnIndex); // getCell uses 0-based row index
          console.log("Summing range: ".concat(rangeAddress, ", placing result in row ").concat(lastDataRow + 1));
          sumCell.formulas = [["=SUM(".concat(rangeAddress, ")")]];
          _context7.n = 7;
          return context.sync();
        case 7:
          sumCell.load(["values", "address"]);
          _context7.n = 8;
          return context.sync();
        case 8:
          return _context7.a(2, {
            success: true,
            message: "".concat(params.columnName, " \uC5F4\uC758 \uD569\uACC4\uB97C ").concat(sumCell.address, "\uC5D0 \uACC4\uC0B0\uD588\uC2B5\uB2C8\uB2E4: ").concat(formatNumber(sumCell.values[0][0])),
            value: sumCell.values[0][0]
          });
        case 9:
          // Original logic for range-based sum
          sourceRange = params.sourceRange ? worksheet.getRange(params.sourceRange) : context.workbook.getSelectedRange();
          sourceRange.load(["columnIndex", "rowIndex", "rowCount", "address"]);
          _context7.n = 10;
          return context.sync();
        case 10:
          if (!(params.addNewRow === true)) {
            _context7.n = 13;
            break;
          }
          // Add new row for sum
          column = sourceRange.columnIndex;
          lastRow = sourceRange.rowIndex + sourceRange.rowCount;
          newCell = worksheet.getCell(lastRow, column);
          newCell.formulas = [["=SUM(".concat(sourceRange.address, ")")]];
          _context7.n = 11;
          return context.sync();
        case 11:
          newCell.load(["values", "address"]);
          _context7.n = 12;
          return context.sync();
        case 12:
          return _context7.a(2, {
            success: true,
            message: "".concat(newCell.address, "\uC5D0 \uD569\uACC4\uB97C \uACC4\uC0B0\uD588\uC2B5\uB2C8\uB2E4: ").concat(formatNumber(newCell.values[0][0])),
            value: newCell.values[0][0]
          });
        case 13:
          if (!params.targetCell) {
            _context7.n = 16;
            break;
          }
          // Sum to specific cell
          targetCell = worksheet.getRange(params.targetCell);
          targetCell.formulas = [["=SUM(".concat(sourceRange.address, ")")]];
          _context7.n = 14;
          return context.sync();
        case 14:
          targetCell.load(["values", "address"]);
          _context7.n = 15;
          return context.sync();
        case 15:
          return _context7.a(2, {
            success: true,
            message: "".concat(targetCell.address, "\uC5D0 \uD569\uACC4\uB97C \uACC4\uC0B0\uD588\uC2B5\uB2C8\uB2E4: ").concat(formatNumber(targetCell.values[0][0])),
            value: targetCell.values[0][0]
          });
        case 16:
          // Default: add sum below the range
          _column = sourceRange.columnIndex;
          _lastRow = sourceRange.rowIndex + sourceRange.rowCount;
          _newCell = worksheet.getCell(_lastRow, _column);
          _newCell.formulas = [["=SUM(".concat(sourceRange.address, ")")]];
          _context7.n = 17;
          return context.sync();
        case 17:
          _newCell.load(["values", "address"]);
          _context7.n = 18;
          return context.sync();
        case 18:
          return _context7.a(2, {
            success: true,
            message: "".concat(_newCell.address, "\uC5D0 \uD569\uACC4\uB97C \uACC4\uC0B0\uD588\uC2B5\uB2C8\uB2E4: ").concat(formatNumber(_newCell.values[0][0])),
            value: _newCell.values[0][0]
          });
        case 19:
          return _context7.a(2);
      }
    }, _callee7);
  }));
  return _executeSum.apply(this, arguments);
}
function executeAverage(_x1, _x10) {
  return _executeAverage.apply(this, arguments);
} // Count cells
function _executeAverage() {
  _executeAverage = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee8(context, params) {
    var worksheet, sourceRange, targetCell, column, lastRow, newCell;
    return _regenerator().w(function (_context8) {
      while (1) switch (_context8.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          sourceRange = params.sourceRange ? worksheet.getRange(params.sourceRange) : context.workbook.getSelectedRange();
          sourceRange.load(["columnIndex", "rowIndex", "rowCount", "address"]);
          _context8.n = 1;
          return context.sync();
        case 1:
          if (!params.targetCell) {
            _context8.n = 4;
            break;
          }
          targetCell = worksheet.getRange(params.targetCell);
          targetCell.formulas = [["=AVERAGE(".concat(sourceRange.address, ")")]];
          _context8.n = 2;
          return context.sync();
        case 2:
          targetCell.load(["values", "address"]);
          _context8.n = 3;
          return context.sync();
        case 3:
          return _context8.a(2, {
            success: true,
            message: "".concat(targetCell.address, "\uC5D0 \uD3C9\uADE0\uC744 \uACC4\uC0B0\uD588\uC2B5\uB2C8\uB2E4: ").concat(formatNumber(targetCell.values[0][0])),
            value: targetCell.values[0][0]
          });
        case 4:
          column = sourceRange.columnIndex;
          lastRow = sourceRange.rowIndex + sourceRange.rowCount;
          newCell = worksheet.getCell(lastRow, column);
          newCell.formulas = [["=AVERAGE(".concat(sourceRange.address, ")")]];
          _context8.n = 5;
          return context.sync();
        case 5:
          newCell.load(["values", "address"]);
          _context8.n = 6;
          return context.sync();
        case 6:
          return _context8.a(2, {
            success: true,
            message: "".concat(newCell.address, "\uC5D0 \uD3C9\uADE0\uC744 \uACC4\uC0B0\uD588\uC2B5\uB2C8\uB2E4: ").concat(formatNumber(newCell.values[0][0])),
            value: newCell.values[0][0]
          });
        case 7:
          return _context8.a(2);
      }
    }, _callee8);
  }));
  return _executeAverage.apply(this, arguments);
}
function executeCount(_x11, _x12) {
  return _executeCount.apply(this, arguments);
} // Format cells
function _executeCount() {
  _executeCount = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee9(context, params) {
    var worksheet, sourceRange, formula, criteria, resultCell, usedRange, lastRow, lastCol;
    return _regenerator().w(function (_context9) {
      while (1) switch (_context9.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          sourceRange = params.sourceRange ? worksheet.getRange(params.sourceRange) : context.workbook.getSelectedRange();
          sourceRange.load(["address"]);
          _context9.n = 1;
          return context.sync();
        case 1:
          if (params.countType === 'countif' && params.condition !== undefined) {
            if (params.operator === 'contains') {
              criteria = "\"*".concat(params.condition, "*\"");
            } else if (params.operator === 'equals') {
              criteria = typeof params.condition === 'string' ? "\"".concat(params.condition, "\"") : params.condition;
            } else if (params.operator && ['>', '<', '>=', '<=', '<>'].includes(params.operator)) {
              criteria = "\"".concat(params.operator).concat(params.condition, "\"");
            } else {
              if (typeof params.condition === 'string') {
                criteria = "\"*".concat(params.condition, "*\"");
              } else {
                criteria = params.condition;
              }
            }
            formula = "=COUNTIF(".concat(sourceRange.address, ", ").concat(criteria, ")");
          } else if (params.countType === 'counta') {
            formula = "=COUNTA(".concat(sourceRange.address, ")");
          } else {
            formula = "=COUNT(".concat(sourceRange.address, ")");
          }
          if (!params.targetCell) {
            _context9.n = 2;
            break;
          }
          resultCell = worksheet.getRange(params.targetCell);
          _context9.n = 4;
          break;
        case 2:
          usedRange = worksheet.getUsedRange();
          _context9.n = 3;
          return context.sync();
        case 3:
          lastRow = usedRange ? usedRange.rowCount : 1;
          lastCol = usedRange ? usedRange.columnCount : 1;
          resultCell = worksheet.getCell(lastRow, lastCol);
        case 4:
          resultCell.formulas = [[formula]];
          _context9.n = 5;
          return context.sync();
        case 5:
          resultCell.load("values");
          _context9.n = 6;
          return context.sync();
        case 6:
          return _context9.a(2, {
            success: true,
            message: params.countType === 'countif' ? "\"".concat(params.condition, "\"\uC744(\uB97C) \uD3EC\uD568\uD558\uB294 \uC140\uC758 \uAC1C\uC218: ").concat(formatNumber(resultCell.values[0][0]), "\uAC1C") : "\uAC1C\uC218\uB97C \uACC4\uC0B0\uD588\uC2B5\uB2C8\uB2E4: ".concat(formatNumber(resultCell.values[0][0]), "\uAC1C"),
            value: resultCell.values[0][0]
          });
      }
    }, _callee9);
  }));
  return _executeCount.apply(this, arguments);
}
function executeFormat(_x13, _x14) {
  return _executeFormat.apply(this, arguments);
} // Sort data
function _executeFormat() {
  _executeFormat = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee0(context, params) {
    var worksheet, range, format;
    return _regenerator().w(function (_context0) {
      while (1) switch (_context0.n) {
        case 0:
          console.log('executeFormat called with params:', params);
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          range = params.range ? worksheet.getRange(params.range) : context.workbook.getSelectedRange();
          range.load("format");
          _context0.n = 1;
          return context.sync();
        case 1:
          if (params.bold !== undefined) {
            range.format.font.bold = params.bold;
          }
          if (params.italic !== undefined) {
            range.format.font.italic = params.italic;
          }
          if (params.fontSize) {
            range.format.font.size = params.fontSize;
          }
          if (params.fontColor) {
            range.format.font.color = params.fontColor;
          }
          if (params.backgroundColor) {
            range.format.fill.color = params.backgroundColor;
          }
          if (params.horizontalAlignment) {
            range.format.horizontalAlignment = params.horizontalAlignment === 'left' ? 'Left' : params.horizontalAlignment === 'center' ? 'Center' : params.horizontalAlignment === 'right' ? 'Right' : 'General';
          }
          if (params.numberFormat) {
            // Handle specific format types
            format = params.numberFormat;
            if (format === 'number' || format === '숫자') {
              format = '#,##0';
            } else if (format === 'currency' || format === '원화' || format === 'won' || format === '통화') {
              format = '₩#,##0';
            } else if (format === 'currency_decimal' || format === '원화_소수') {
              format = '₩#,##0.00';
            } else if (format === 'percentage' || format === '퍼센트' || format === '백분율') {
              format = '0%';
            } else if (format === 'percentage_decimal' || format === '퍼센트_소수') {
              format = '0.00%';
            } else if (format === 'date' || format === '날짜') {
              format = 'yyyy-mm-dd';
            } else if (format === 'time' || format === '시간') {
              format = 'hh:mm:ss';
            } else if (format === 'text' || format === '텍스트') {
              format = '@';
            } else if (format === 'general' || format === '일반') {
              format = 'General';
            }

            // Set number format for the range
            range.numberFormat = format;
          }
          _context0.n = 2;
          return context.sync();
        case 2:
          return _context0.a(2, {
            success: true,
            message: '서식이 적용되었습니다.'
          });
      }
    }, _callee0);
  }));
  return _executeFormat.apply(this, arguments);
}
function executeSort(_x15, _x16) {
  return _executeSort.apply(this, arguments);
} // Create chart
function _executeSort() {
  _executeSort = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee1(context, params) {
    var worksheet, range, column, ascending, columnLetter;
    return _regenerator().w(function (_context1) {
      while (1) switch (_context1.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          range = params.range ? worksheet.getRange(params.range) : worksheet.getUsedRange();
          column = params.column || 1;
          ascending = params.ascending !== undefined ? params.ascending : true;
          range.sort.apply([{
            key: column - 1,
            // Excel API uses 0-based index
            ascending: ascending
          }]);
          _context1.n = 1;
          return context.sync();
        case 1:
          columnLetter = getColumnLetter(column - 1);
          return _context1.a(2, {
            success: true,
            message: "".concat(columnLetter, "\uC5F4 \uAE30\uC900\uC73C\uB85C ").concat(ascending ? '오름차순' : '내림차순', " \uC815\uB82C\uB418\uC5C8\uC2B5\uB2C8\uB2E4.")
          });
      }
    }, _callee1);
  }));
  return _executeSort.apply(this, arguments);
}
function executeChart(_x17, _x18) {
  return _executeChart.apply(this, arguments);
} // Add conditional formatting
function _executeChart() {
  _executeChart = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee10(context, params) {
    var worksheet, sourceData, chartType, chart;
    return _regenerator().w(function (_context10) {
      while (1) switch (_context10.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          sourceData = params.range ? worksheet.getRange(params.range) : context.workbook.getSelectedRange();
          console.log('Creating chart with params:', params);

          // Load the source data to ensure it's available
          sourceData.load(["address", "values"]);
          _context10.n = 1;
          return context.sync();
        case 1:
          console.log('Chart data range:', sourceData.address);

          // Determine chart type

          if (params.chartType === 'bar' || params.chartType === '막대') {
            chartType = Excel.ChartType.columnClustered;
          } else if (params.chartType === 'line' || params.chartType === '선') {
            chartType = Excel.ChartType.line;
          } else if (params.chartType === 'pie' || params.chartType === '원') {
            chartType = Excel.ChartType.pie;
          } else if (params.chartType === 'scatter' || params.chartType === '분산형') {
            chartType = Excel.ChartType.xyscatter;
          } else {
            // Default to column chart
            chartType = Excel.ChartType.columnClustered;
          }

          // Create the chart
          chart = worksheet.charts.add(chartType, sourceData, Excel.ChartSeriesBy.auto); // Set chart properties
          chart.title.text = params.title || '차트';
          chart.height = 300;
          chart.width = 400;

          // Position the chart
          chart.left = params.offsetX || 100;
          chart.top = params.offsetY || 100;

          // Set legend position
          chart.legend.position = Excel.ChartLegendPosition.bottom;
          chart.legend.visible = true;
          _context10.n = 2;
          return context.sync();
        case 2:
          return _context10.a(2, {
            success: true,
            message: "".concat(sourceData.address, " \uBC94\uC704\uB85C ").concat(params.chartType || '막대', " \uCC28\uD2B8\uAC00 \uC0DD\uC131\uB418\uC5C8\uC2B5\uB2C8\uB2E4.")
          });
      }
    }, _callee10);
  }));
  return _executeChart.apply(this, arguments);
}
function executeConditionalFormat(_x19, _x20) {
  return _executeConditionalFormat.apply(this, arguments);
} // Translate column contents
function _executeConditionalFormat() {
  _executeConditionalFormat = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee11(context, params) {
    var worksheet, range, conditionalFormat;
    return _regenerator().w(function (_context11) {
      while (1) switch (_context11.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          range = params.range ? worksheet.getRange(params.range) : worksheet.getUsedRange();
          console.log('Applying conditional format with params:', params);

          // Simply apply conditional format to the entire range
          // Excel will automatically skip non-numeric cells for numeric comparisons
          conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue); // Set the rule based on condition
          if (params.condition === 'greater_than' && params.value !== undefined) {
            conditionalFormat.cellValue.rule = {
              formula1: params.value.toString(),
              operator: Excel.ConditionalCellValueOperator.greaterThan
            };
          } else if (params.condition === 'less_than' && params.value !== undefined) {
            conditionalFormat.cellValue.rule = {
              formula1: params.value.toString(),
              operator: Excel.ConditionalCellValueOperator.lessThan
            };
          } else if (params.condition === 'equal_to' && params.value !== undefined) {
            conditionalFormat.cellValue.rule = {
              formula1: params.value.toString(),
              operator: Excel.ConditionalCellValueOperator.equalTo
            };
          } else if (params.condition === 'text_contains' && params.value !== undefined) {
            conditionalFormat.cellValue.rule = {
              formula1: params.value.toString(),
              operator: Excel.ConditionalCellValueOperator.containsText
            };
          } else {
            // Default to greater than
            conditionalFormat.cellValue.rule = {
              formula1: (params.value || 0).toString(),
              operator: Excel.ConditionalCellValueOperator.greaterThan
            };
          }

          // Set the format
          conditionalFormat.cellValue.format.fill.color = params.backgroundColor || "#00FF00";
          if (params.fontColor) {
            conditionalFormat.cellValue.format.font.color = params.fontColor;
          }
          if (params.bold) {
            conditionalFormat.cellValue.format.font.bold = true;
          }
          _context11.n = 1;
          return context.sync();
        case 1:
          return _context11.a(2, {
            success: true,
            message: "\uC870\uAC74\uBD80 \uC11C\uC2DD\uC774 \uC801\uC6A9\uB418\uC5C8\uC2B5\uB2C8\uB2E4. (".concat(params.condition, " ").concat(params.value || '', ")")
          });
      }
    }, _callee11);
  }));
  return _executeConditionalFormat.apply(this, arguments);
}
function executeTranslate(_x21, _x22) {
  return _executeTranslate.apply(this, arguments);
} // Translate a batch of texts through proxy
function _executeTranslate() {
  _executeTranslate = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee12(context, params) {
    var worksheet, sourceRange, targetColumnIndex, usedRange, columnLetter, columnIndex, targetColumnLetter, _columnIndex, _usedRange, sourceValues, translations, batchSize, i, batch, batchTexts, translatedBatch, j, translationIndex, _j, translatedText, _j2, targetRange, stringTranslations, isEmpty, _i, cellRow, cellCol, cell, cellValue, headerCell, sourceHeaderCell;
    return _regenerator().w(function (_context12) {
      while (1) switch (_context12.n) {
        case 0:
          console.log('executeTranslate called with params:', params);
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          if (params.targetLanguage) {
            _context12.n = 1;
            break;
          }
          return _context12.a(2, {
            success: false,
            error: '대상 언어를 지정해주세요.'
          });
        case 1:
          if (!params.sourceRange) {
            _context12.n = 8;
            break;
          }
          if (!params.sourceRange.match(/^[A-Z]+:[A-Z]+$/)) {
            _context12.n = 6;
            break;
          }
          console.log('Handling column range:', params.sourceRange);
          usedRange = worksheet.getUsedRange();
          if (usedRange) {
            _context12.n = 2;
            break;
          }
          return _context12.a(2, {
            success: false,
            error: '번역할 데이터가 없습니다.'
          });
        case 2:
          usedRange.load(["rowCount", "columnCount"]);
          _context12.n = 3;
          return context.sync();
        case 3:
          console.log('Used range:', {
            rowCount: usedRange.rowCount,
            columnCount: usedRange.columnCount
          });

          // Get the column letter from the range (e.g., "C" from "C:C")
          columnLetter = params.sourceRange.split(':')[0];
          columnIndex = columnLetter.charCodeAt(0) - 65; // Get only the used portion of the column, starting from row 2 (skip header)
          if (!(usedRange.rowCount > 1)) {
            _context12.n = 4;
            break;
          }
          sourceRange = worksheet.getRangeByIndexes(1, columnIndex, usedRange.rowCount - 1, 1);
          _context12.n = 5;
          break;
        case 4:
          return _context12.a(2, {
            success: false,
            error: '번역할 데이터가 없습니다.'
          });
        case 5:
          // Handle target range similarly
          if (params.targetRange && params.targetRange.match(/^[A-Z]+:[A-Z]+$/)) {
            targetColumnLetter = params.targetRange.split(':')[0];
            targetColumnIndex = targetColumnLetter.charCodeAt(0) - 65;
            console.log('Target column calculation:', {
              targetRange: params.targetRange,
              targetColumnLetter: targetColumnLetter,
              targetColumnIndex: targetColumnIndex
            });
          } else {
            // If no target specified, use next column
            targetColumnIndex = columnIndex + 1;
            console.log('Using next column as target:', {
              sourceColumnIndex: columnIndex,
              targetColumnIndex: targetColumnIndex
            });
          }
          _context12.n = 7;
          break;
        case 6:
          sourceRange = worksheet.getRange(params.sourceRange);
        case 7:
          _context12.n = 13;
          break;
        case 8:
          if (!params.sourceColumn) {
            _context12.n = 12;
            break;
          }
          _columnIndex = params.sourceColumn.charCodeAt(0) - 65;
          _usedRange = worksheet.getUsedRange();
          _usedRange.load(["rowCount"]);
          _context12.n = 9;
          return context.sync();
        case 9:
          if (!(_usedRange.rowCount > 1)) {
            _context12.n = 10;
            break;
          }
          sourceRange = worksheet.getRangeByIndexes(1, _columnIndex, _usedRange.rowCount - 1, 1);
          _context12.n = 11;
          break;
        case 10:
          return _context12.a(2, {
            success: false,
            error: '번역할 데이터가 없습니다.'
          });
        case 11:
          targetColumnIndex = params.targetColumn === 'next' ? _columnIndex + 1 : params.targetColumn ? params.targetColumn.charCodeAt(0) - 65 : _columnIndex + 1;
          _context12.n = 13;
          break;
        case 12:
          return _context12.a(2, {
            success: false,
            error: '번역할 열을 지정해주세요.'
          });
        case 13:
          sourceRange.load(["values", "rowIndex", "columnIndex", "rowCount"]);
          _context12.n = 14;
          return context.sync();
        case 14:
          console.log('Source range loaded:', {
            rowIndex: sourceRange.rowIndex,
            columnIndex: sourceRange.columnIndex,
            rowCount: sourceRange.rowCount,
            values: sourceRange.values ? "".concat(sourceRange.values.length, " rows") : 'null'
          });
          sourceValues = sourceRange.values;
          if (!(!sourceValues || sourceValues.length === 0)) {
            _context12.n = 15;
            break;
          }
          return _context12.a(2, {
            success: false,
            error: '번역할 데이터가 없습니다.'
          });
        case 15:
          translations = [];
          batchSize = 20; // Translate in batches
          i = 0;
        case 16:
          if (!(i < sourceValues.length)) {
            _context12.n = 21;
            break;
          }
          batch = sourceValues.slice(i, Math.min(i + batchSize, sourceValues.length));
          batchTexts = batch.map(function (row) {
            return row[0];
          }).filter(function (text) {
            return text;
          });
          if (!(batchTexts.length > 0)) {
            _context12.n = 18;
            break;
          }
          _context12.n = 17;
          return translateBatch(batchTexts, params.targetLanguage, params.sourceLanguage);
        case 17:
          translatedBatch = _context12.v;
          // Check if translatedBatch is valid
          if (!translatedBatch || !Array.isArray(translatedBatch)) {
            console.error('Invalid translation batch received:', translatedBatch);
            // Fill with empty strings if translation failed
            for (j = 0; j < batch.length; j++) {
              translations.push(['']);
            }
          } else {
            console.log('Processing translation batch:', {
              batchLength: batch.length,
              translatedBatchLength: translatedBatch.length,
              sampleTranslations: translatedBatch.slice(0, 3),
              firstTranslation: translatedBatch[0],
              translationType: _typeof(translatedBatch[0]),
              rawData: JSON.stringify(translatedBatch.slice(0, 3))
            });
            translationIndex = 0;
            for (_j = 0; _j < batch.length; _j++) {
              if (batch[_j][0]) {
                translatedText = translatedBatch[translationIndex] || '';
                translations.push([translatedText]);
                if (_j < 3) {
                  console.log("Translation ".concat(_j, ": \"").concat(batch[_j][0], "\" -> \"").concat(translatedText, "\""));
                }
                translationIndex++;
              } else {
                translations.push(['']);
              }
            }
          }
          _context12.n = 19;
          break;
        case 18:
          for (_j2 = 0; _j2 < batch.length; _j2++) {
            translations.push(['']);
          }
        case 19:
          // Show progress
          if (i % 100 === 0 && i > 0) {
            showStatus("\uBC88\uC5ED \uC911... ".concat(Math.round(i / sourceValues.length * 100), "%"), 'info');
          }
        case 20:
          i += batchSize;
          _context12.n = 16;
          break;
        case 21:
          // Write translations
          console.log('Writing translations to target column:', {
            rowIndex: sourceRange.rowIndex,
            targetColumnIndex: targetColumnIndex || sourceRange.columnIndex + 1,
            translationsCount: translations.length,
            sampleTranslations: translations.slice(0, 3).map(function (t) {
              return t[0];
            })
          });
          console.log('First 5 translations raw:', JSON.stringify(translations.slice(0, 5)));
          console.log('Translation content check:', {
            first: translations[0] ? translations[0][0] : 'null',
            second: translations[1] ? translations[1][0] : 'null',
            third: translations[2] ? translations[2][0] : 'null',
            isEmpty: translations[0] && translations[0][0] === ''
          });
          targetRange = worksheet.getRangeByIndexes(sourceRange.rowIndex, targetColumnIndex || sourceRange.columnIndex + 1, translations.length, 1);
          targetRange.load(["address", "values"]);
          _context12.n = 22;
          return context.sync();
        case 22:
          console.log('Target range address:', targetRange.address);
          console.log('Existing target values (first 3):', targetRange.values.slice(0, 3));

          // Clear existing values first
          targetRange.clear(Excel.ClearApplyTo.contents);
          _context12.n = 23;
          return context.sync();
        case 23:
          console.log('Target range cleared');

          // Ensure translations are properly formatted as 2D array
          console.log('Setting target range values:', {
            translationsLength: translations.length,
            firstTranslation: translations[0],
            isArray: Array.isArray(translations),
            is2DArray: Array.isArray(translations[0])
          });

          // Try setting values with explicit string conversion
          stringTranslations = translations.map(function (row) {
            return [String(row[0] || '')];
          });
          console.log('String translations (first 3):', stringTranslations.slice(0, 3));

          // Try using numberFormat to ensure cells are treated as text
          targetRange.numberFormat = [["@"]]; // @ means text format
          _context12.n = 24;
          return context.sync();
        case 24:
          targetRange.values = stringTranslations;

          // Force Excel to update
          targetRange.format.autofitColumns();
          targetRange.format.font.color = "#000000"; // Ensure text is visible
          targetRange.format.fill.color = "#FFFFFF"; // White background
          _context12.n = 25;
          return context.sync();
        case 25:
          console.log('Translations written to Excel');

          // Verify the values were actually written
          targetRange.load(["values", "text", "valueTypes"]);
          _context12.n = 26;
          return context.sync();
        case 26:
          console.log('Verification - Target range:', {
            address: targetRange.address,
            values: targetRange.values.slice(0, 3),
            text: targetRange.text.slice(0, 3),
            valueTypes: targetRange.valueTypes.slice(0, 3),
            actualFirstValue: targetRange.values[0] ? targetRange.values[0][0] : 'null',
            firstThreeValues: [targetRange.values[0] ? targetRange.values[0][0] : 'empty', targetRange.values[1] ? targetRange.values[1][0] : 'empty', targetRange.values[2] ? targetRange.values[2][0] : 'empty']
          });

          // Try alternative method - set each cell individually for debugging
          if (!(translations.length > 0)) {
            _context12.n = 31;
            break;
          }
          isEmpty = !targetRange.values[0] || !targetRange.values[0][0] || targetRange.values[0][0] === '';
          console.log('Checking if values are empty:', {
            isEmpty: isEmpty,
            firstValue: targetRange.values[0] ? targetRange.values[0][0] : 'null',
            firstValueLength: targetRange.values[0] && targetRange.values[0][0] ? targetRange.values[0][0].length : 0
          });
          if (!isEmpty) {
            _context12.n = 31;
            break;
          }
          console.log('Values not visible, trying individual cell approach...');
          // Try setting just the first few cells individually
          _i = 0;
        case 27:
          if (!(_i < Math.min(3, translations.length))) {
            _context12.n = 31;
            break;
          }
          cellRow = sourceRange.rowIndex + _i;
          cellCol = targetColumnIndex || sourceRange.columnIndex + 1;
          cell = worksheet.getCell(cellRow, cellCol);
          cellValue = stringTranslations[_i][0];
          console.log("Setting cell (".concat(cellRow, ",").concat(cellCol, ") to: \"").concat(cellValue, "\""));
          cell.values = [[cellValue]];
          cell.format.font.color = "#000000";
          cell.format.fill.color = "#FFFF00"; // Yellow background to make it visible
          _context12.n = 28;
          return context.sync();
        case 28:
          // Verify it was set
          cell.load("values");
          _context12.n = 29;
          return context.sync();
        case 29:
          console.log("Cell ".concat(_i, " after setting:"), cell.values[0][0]);
        case 30:
          _i++;
          _context12.n = 27;
          break;
        case 31:
          // Add header
          headerCell = worksheet.getCell(0, targetColumnIndex || sourceRange.columnIndex + 1);
          sourceHeaderCell = worksheet.getCell(0, sourceRange.columnIndex);
          sourceHeaderCell.load("values");
          _context12.n = 32;
          return context.sync();
        case 32:
          headerCell.values = [["".concat(sourceHeaderCell.values[0][0], " (").concat(params.targetLanguage, ")")]];
          _context12.n = 33;
          return context.sync();
        case 33:
          return _context12.a(2, {
            success: true,
            message: "\uBC88\uC5ED\uC774 \uC644\uB8CC\uB418\uC5C8\uC2B5\uB2C8\uB2E4. (".concat(sourceValues.length, "\uAC1C \uD56D\uBAA9)")
          });
      }
    }, _callee12);
  }));
  return _executeTranslate.apply(this, arguments);
}
function translateBatch(_x23, _x24, _x25) {
  return _translateBatch.apply(this, arguments);
} // Helper functions
function _translateBatch() {
  _translateBatch = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee13(texts, targetLanguage, sourceLanguage) {
    var response, result, _t7;
    return _regenerator().w(function (_context13) {
      while (1) switch (_context13.p = _context13.n) {
        case 0:
          _context13.p = 0;
          console.log('translateBatch called with', texts.length, 'texts');
          _context13.n = 1;
          return fetch(API_PROXY_URL, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              command: "Translate these items to ".concat(targetLanguage, ": ").concat(texts.join(', ')),
              sheetContext: {
                operation: 'translate_batch',
                texts: texts,
                targetLanguage: targetLanguage,
                sourceLanguage: sourceLanguage
              }
            })
          });
        case 1:
          response = _context13.v;
          console.log('Translation response status:', response.status);
          if (response.ok) {
            _context13.n = 2;
            break;
          }
          throw new Error("HTTP error! status: ".concat(response.status));
        case 2:
          _context13.n = 3;
          return response.json();
        case 3:
          result = _context13.v;
          console.log('Translation result:', result);
          if (!(result.success && result.data && result.data.translations)) {
            _context13.n = 4;
            break;
          }
          console.log('Translations received:', result.data.translations.length);
          return _context13.a(2, result.data.translations);
        case 4:
          if (!result.error) {
            _context13.n = 5;
            break;
          }
          console.error('Translation API error:', result.error);
          throw new Error(result.error);
        case 5:
          console.error('Invalid translation response structure:', result);
          throw new Error('번역 응답을 받을 수 없습니다.');
        case 6:
          _context13.n = 8;
          break;
        case 7:
          _context13.p = 7;
          _t7 = _context13.v;
          console.error('Translation error:', _t7);
          return _context13.a(2, texts.map(function () {
            return '';
          }));
        case 8:
          return _context13.a(2);
      }
    }, _callee13, null, [[0, 7]]);
  }));
  return _translateBatch.apply(this, arguments);
}
function getColumnLetter(columnIndex) {
  var columnLetter = '';
  var tempIndex = columnIndex;
  while (tempIndex >= 0) {
    columnLetter = String.fromCharCode(tempIndex % 26 + 65) + columnLetter;
    tempIndex = Math.floor(tempIndex / 26) - 1;
  }
  return columnLetter;
}
function formatNumber(value) {
  if (typeof value === 'number') {
    if (Number.isInteger(value)) {
      return value.toLocaleString('ko-KR');
    } else {
      return value.toLocaleString('ko-KR', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
      });
    }
  }
  return value;
}
function showStatus(message, type) {
  var status = document.getElementById('status');
  status.textContent = message;
  status.className = 'status-message ' + type;
  status.style.display = 'block';
  if (window.statusTimeout) {
    clearTimeout(window.statusTimeout);
  }
  if (type === 'success') {
    window.statusTimeout = setTimeout(function () {
      status.style.display = 'none';
    }, 5000);
  }
}
function clearInput() {
  document.getElementById('commandInput').value = '';
  document.getElementById('status').style.display = 'none';
}
function setCommand(command) {
  document.getElementById('commandInput').value = command;
  document.getElementById('commandInput').focus();
}
function showSettings() {
  showStatus('이 애드인은 보안 서버를 통해 AI 기능을 제공합니다. 별도의 API 키 설정이 필요하지 않습니다.', 'info');
}

// Test backend connection
function testBackendConnection() {
  return _testBackendConnection.apply(this, arguments);
} // Additional operations (compress, filter, insert, delete, formula, retry_translation)
// These would need to be implemented based on Excel JavaScript API capabilities
function _testBackendConnection() {
  _testBackendConnection = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee14() {
    var testUrl, response, data, _t8;
    return _regenerator().w(function (_context14) {
      while (1) switch (_context14.p = _context14.n) {
        case 0:
          _context14.p = 0;
          console.log('Testing backend connection...');
          testUrl = 'https://excel-addon-backend.vercel.app/api/test';
          _context14.n = 1;
          return fetch(testUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              test: true
            })
          });
        case 1:
          response = _context14.v;
          console.log('Test response status:', response.status);
          console.log('Test response headers:', response.headers);
          _context14.n = 2;
          return response.json();
        case 2:
          data = _context14.v;
          console.log('Test response data:', data);
          showStatus('백엔드 연결 테스트 성공', 'success');
          _context14.n = 4;
          break;
        case 3:
          _context14.p = 3;
          _t8 = _context14.v;
          console.error('Backend test error:', _t8);
          showStatus('백엔드 연결 테스트 실패: ' + _t8.message, 'error');
        case 4:
          return _context14.a(2);
      }
    }, _callee14, null, [[0, 3]]);
  }));
  return _testBackendConnection.apply(this, arguments);
}
function executeCompress(_x26, _x27) {
  return _executeCompress.apply(this, arguments);
}
function _executeCompress() {
  _executeCompress = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee15(context, params) {
    return _regenerator().w(function (_context15) {
      while (1) switch (_context15.n) {
        case 0:
          return _context15.a(2, {
            success: false,
            error: '이 기능은 현재 구현 중입니다.'
          });
      }
    }, _callee15);
  }));
  return _executeCompress.apply(this, arguments);
}
function executeFilter(_x28, _x29) {
  return _executeFilter.apply(this, arguments);
}
function _executeFilter() {
  _executeFilter = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee16(context, params) {
    var worksheet, range;
    return _regenerator().w(function (_context16) {
      while (1) switch (_context16.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          range = params.range ? worksheet.getRange(params.range) : worksheet.getUsedRange(); // Apply autofilter
          range.worksheet.autoFilter.apply(range);
          _context16.n = 1;
          return context.sync();
        case 1:
          return _context16.a(2, {
            success: true,
            message: '필터가 적용되었습니다.'
          });
      }
    }, _callee16);
  }));
  return _executeFilter.apply(this, arguments);
}
function executeInsert(_x30, _x31) {
  return _executeInsert.apply(this, arguments);
}
function _executeInsert() {
  _executeInsert = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee17(context, params) {
    var worksheet, type, position, count, range, _range;
    return _regenerator().w(function (_context17) {
      while (1) switch (_context17.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          type = params.type || 'row';
          position = params.position || 1;
          count = params.count || 1;
          if (type === 'row') {
            range = worksheet.getRangeByIndexes(position - 1, 0, count, 1);
            range.insert(Excel.InsertShiftDirection.down);
          } else {
            _range = worksheet.getRangeByIndexes(0, position - 1, 1, count);
            _range.insert(Excel.InsertShiftDirection.right);
          }
          _context17.n = 1;
          return context.sync();
        case 1:
          return _context17.a(2, {
            success: true,
            message: "".concat(count, "\uAC1C\uC758 ").concat(type === 'row' ? '행' : '열', "\uC774 \uC0BD\uC785\uB418\uC5C8\uC2B5\uB2C8\uB2E4.")
          });
      }
    }, _callee17);
  }));
  return _executeInsert.apply(this, arguments);
}
function executeDelete(_x32, _x33) {
  return _executeDelete.apply(this, arguments);
}
function _executeDelete() {
  _executeDelete = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee18(context, params) {
    var worksheet, type, position, count, range, _range2;
    return _regenerator().w(function (_context18) {
      while (1) switch (_context18.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          type = params.type || 'row';
          position = params.position || 1;
          count = params.count || 1;
          if (type === 'row') {
            range = worksheet.getRangeByIndexes(position - 1, 0, count, 1);
            range.delete(Excel.DeleteShiftDirection.up);
          } else {
            _range2 = worksheet.getRangeByIndexes(0, position - 1, 1, count);
            _range2.delete(Excel.DeleteShiftDirection.left);
          }
          _context18.n = 1;
          return context.sync();
        case 1:
          return _context18.a(2, {
            success: true,
            message: "".concat(count, "\uAC1C\uC758 ").concat(type === 'row' ? '행' : '열', "\uC774 \uC0AD\uC81C\uB418\uC5C8\uC2B5\uB2C8\uB2E4.")
          });
      }
    }, _callee18);
  }));
  return _executeDelete.apply(this, arguments);
}
function executeFormula(_x34, _x35) {
  return _executeFormula.apply(this, arguments);
}
function _executeFormula() {
  _executeFormula = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee19(context, params) {
    var worksheet, range;
    return _regenerator().w(function (_context19) {
      while (1) switch (_context19.n) {
        case 0:
          worksheet = context.workbook.worksheets.getActiveWorksheet();
          range = params.range ? worksheet.getRange(params.range) : context.workbook.getSelectedRange();
          if (params.formula) {
            _context19.n = 1;
            break;
          }
          return _context19.a(2, {
            success: false,
            error: '수식을 지정해주세요.'
          });
        case 1:
          range.formulas = [[params.formula]];
          _context19.n = 2;
          return context.sync();
        case 2:
          return _context19.a(2, {
            success: true,
            message: '수식이 적용되었습니다.'
          });
      }
    }, _callee19);
  }));
  return _executeFormula.apply(this, arguments);
}
function executeRetryTranslation(_x36, _x37) {
  return _executeRetryTranslation.apply(this, arguments);
}
function _executeRetryTranslation() {
  _executeRetryTranslation = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee20(context, params) {
    return _regenerator().w(function (_context20) {
      while (1) switch (_context20.n) {
        case 0:
          return _context20.a(2, {
            success: false,
            error: '이 기능은 현재 구현 중입니다.'
          });
      }
    }, _callee20);
  }));
  return _executeRetryTranslation.apply(this, arguments);
}
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map