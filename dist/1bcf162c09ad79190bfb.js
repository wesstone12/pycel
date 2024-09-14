/*! For license information please see 1bcf162c09ad79190bfb.js.LICENSE.txt */
function _typeof(t){return _typeof="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},_typeof(t)}function _regeneratorRuntime(){"use strict";_regeneratorRuntime=function(){return e};var t,e={},r=Object.prototype,n=r.hasOwnProperty,o=Object.defineProperty||function(t,e,r){t[e]=r.value},i="function"==typeof Symbol?Symbol:{},a=i.iterator||"@@iterator",c=i.asyncIterator||"@@asyncIterator",l=i.toStringTag||"@@toStringTag";function u(t,e,r){return Object.defineProperty(t,e,{value:r,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{u({},"")}catch(t){u=function(t,e,r){return t[e]=r}}function s(t,e,r,n){var i=e&&e.prototype instanceof m?e:m,a=Object.create(i.prototype),c=new G(n||[]);return o(a,"_invoke",{value:S(t,r,c)}),a}function f(t,e,r){try{return{type:"normal",arg:t.call(e,r)}}catch(t){return{type:"throw",arg:t}}}e.wrap=s;var p="suspendedStart",h="suspendedYield",y="executing",d="completed",g={};function m(){}function v(){}function x(){}var b={};u(b,a,(function(){return this}));var w=Object.getPrototypeOf,E=w&&w(w(I([])));E&&E!==r&&n.call(E,a)&&(b=E);var _=x.prototype=m.prototype=Object.create(b);function L(t){["next","throw","return"].forEach((function(e){u(t,e,(function(t){return this._invoke(e,t)}))}))}function O(t,e){function r(o,i,a,c){var l=f(t[o],t,i);if("throw"!==l.type){var u=l.arg,s=u.value;return s&&"object"==_typeof(s)&&n.call(s,"__await")?e.resolve(s.__await).then((function(t){r("next",t,a,c)}),(function(t){r("throw",t,a,c)})):e.resolve(s).then((function(t){u.value=t,a(u)}),(function(t){return r("throw",t,a,c)}))}c(l.arg)}var i;o(this,"_invoke",{value:function(t,n){function o(){return new e((function(e,o){r(t,n,e,o)}))}return i=i?i.then(o,o):o()}})}function S(e,r,n){var o=p;return function(i,a){if(o===y)throw Error("Generator is already running");if(o===d){if("throw"===i)throw a;return{value:t,done:!0}}for(n.method=i,n.arg=a;;){var c=n.delegate;if(c){var l=P(c,n);if(l){if(l===g)continue;return l}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(o===p)throw o=d,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);o=y;var u=f(e,r,n);if("normal"===u.type){if(o=n.done?d:h,u.arg===g)continue;return{value:u.arg,done:n.done}}"throw"===u.type&&(o=d,n.method="throw",n.arg=u.arg)}}}function P(e,r){var n=r.method,o=e.iterator[n];if(o===t)return r.delegate=null,"throw"===n&&e.iterator.return&&(r.method="return",r.arg=t,P(e,r),"throw"===r.method)||"return"!==n&&(r.method="throw",r.arg=new TypeError("The iterator does not provide a '"+n+"' method")),g;var i=f(o,e.iterator,r.arg);if("throw"===i.type)return r.method="throw",r.arg=i.arg,r.delegate=null,g;var a=i.arg;return a?a.done?(r[e.resultName]=a.value,r.next=e.nextLoc,"return"!==r.method&&(r.method="next",r.arg=t),r.delegate=null,g):a:(r.method="throw",r.arg=new TypeError("iterator result is not an object"),r.delegate=null,g)}function j(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function k(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function G(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(j,this),this.reset(!0)}function I(e){if(e||""===e){var r=e[a];if(r)return r.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,i=function r(){for(;++o<e.length;)if(n.call(e,o))return r.value=e[o],r.done=!1,r;return r.value=t,r.done=!0,r};return i.next=i}}throw new TypeError(_typeof(e)+" is not iterable")}return v.prototype=x,o(_,"constructor",{value:x,configurable:!0}),o(x,"constructor",{value:v,configurable:!0}),v.displayName=u(x,l,"GeneratorFunction"),e.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===v||"GeneratorFunction"===(e.displayName||e.name))},e.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,x):(t.__proto__=x,u(t,l,"GeneratorFunction")),t.prototype=Object.create(_),t},e.awrap=function(t){return{__await:t}},L(O.prototype),u(O.prototype,c,(function(){return this})),e.AsyncIterator=O,e.async=function(t,r,n,o,i){void 0===i&&(i=Promise);var a=new O(s(t,r,n,o),i);return e.isGeneratorFunction(r)?a:a.next().then((function(t){return t.done?t.value:a.next()}))},L(_),u(_,l,"Generator"),u(_,a,(function(){return this})),u(_,"toString",(function(){return"[object Generator]"})),e.keys=function(t){var e=Object(t),r=[];for(var n in e)r.push(n);return r.reverse(),function t(){for(;r.length;){var n=r.pop();if(n in e)return t.value=n,t.done=!1,t}return t.done=!0,t}},e.values=I,G.prototype={constructor:G,reset:function(e){if(this.prev=0,this.next=0,this.sent=this._sent=t,this.done=!1,this.delegate=null,this.method="next",this.arg=t,this.tryEntries.forEach(k),!e)for(var r in this)"t"===r.charAt(0)&&n.call(this,r)&&!isNaN(+r.slice(1))&&(this[r]=t)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(e){if(this.done)throw e;var r=this;function o(n,o){return c.type="throw",c.arg=e,r.next=n,o&&(r.method="next",r.arg=t),!!o}for(var i=this.tryEntries.length-1;i>=0;--i){var a=this.tryEntries[i],c=a.completion;if("root"===a.tryLoc)return o("end");if(a.tryLoc<=this.prev){var l=n.call(a,"catchLoc"),u=n.call(a,"finallyLoc");if(l&&u){if(this.prev<a.catchLoc)return o(a.catchLoc,!0);if(this.prev<a.finallyLoc)return o(a.finallyLoc)}else if(l){if(this.prev<a.catchLoc)return o(a.catchLoc,!0)}else{if(!u)throw Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return o(a.finallyLoc)}}}},abrupt:function(t,e){for(var r=this.tryEntries.length-1;r>=0;--r){var o=this.tryEntries[r];if(o.tryLoc<=this.prev&&n.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var i=o;break}}i&&("break"===t||"continue"===t)&&i.tryLoc<=e&&e<=i.finallyLoc&&(i=null);var a=i?i.completion:{};return a.type=t,a.arg=e,i?(this.method="next",this.next=i.finallyLoc,g):this.complete(a)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),g},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.finallyLoc===t)return this.complete(r.completion,r.afterLoc),k(r),g}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.tryLoc===t){var n=r.completion;if("throw"===n.type){var o=n.arg;k(r)}return o}}throw Error("illegal catch attempt")},delegateYield:function(e,r,n){return this.delegate={iterator:I(e),resultName:r,nextLoc:n},"next"===this.method&&(this.arg=t),g}},e}function asyncGeneratorStep(t,e,r,n,o,i,a){try{var c=t[i](a),l=c.value}catch(t){return void r(t)}c.done?e(l):Promise.resolve(l).then(n,o)}function _asyncToGenerator(t){return function(){var e=this,r=arguments;return new Promise((function(n,o){var i=t.apply(e,r);function a(t){asyncGeneratorStep(i,n,o,a,c,"next",t)}function c(t){asyncGeneratorStep(i,n,o,a,c,"throw",t)}a(void 0)}))}}Office.onReady((function(t){if(console.log("Office.onReady called with info:",t),t.host===Office.HostType.Excel){console.log("Host is Excel. Proceeding to display app body."),document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex";var e=document.getElementById("execute-script");e?e.onclick=executeScript:console.warn("'execute-script' button not found in the DOM.")}else console.log("Host is not Excel. Add-in not initialized.")}));export function executeScript(){return _executeScript.apply(this,arguments)}function _executeScript(){return(_executeScript=_asyncToGenerator(_regeneratorRuntime().mark((function t(){var e,r,n,o,i,a;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(t.prev=0,(e=document.getElementById("script-input").value).trim()){t.next=5;break}return displayErrorDialog("Please enter a Python script to execute."),t.abrupt("return");case 5:return console.log("Sending script to Python server."),(r=document.getElementById("loading-indicator"))&&(r.style.display="block"),t.next=10,fetch("http://127.0.0.1:8000/execute-script",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({script:e})});case 10:if(n=t.sent,console.log("Received response status:",n.status),n.ok){t.next=17;break}return t.next=15,n.json();case 15:throw o=t.sent,new Error("Server error: ".concat(o.detail));case 17:return t.next=19,n.json();case 19:i=t.sent,console.log("Received plot data:",i.plot),r&&(r.style.display="none"),displayPlot(i.plot),t.next=31;break;case 25:t.prev=25,t.t0=t.catch(0),console.error("Error executing script:",t.t0),(a=document.getElementById("loading-indicator"))&&(a.style.display="none"),displayErrorDialog("Failed to execute the Python script.");case 31:case"end":return t.stop()}}),t,null,[[0,25]])})))).apply(this,arguments)}function displayPlot(t){var e=document.getElementById("plot-image");e?(e.src="data:image/png;base64,".concat(t),console.log("Plot image source updated.")):(console.error("'plot-image' element not found in the DOM."),displayErrorDialog("Failed to display the plot image."))}function displayErrorDialog(t){var e='\n    <!DOCTYPE html>\n    <html>\n      <head>\n        <title>Error</title>\n        <style>\n          body { font-family: Arial, sans-serif; padding: 20px; }\n          .error { color: red; }\n          button { margin-top: 20px; padding: 10px 20px; }\n        </style>\n      </head>\n      <body>\n        <h2 class="error">Error</h2>\n        <p>'.concat(t,"</p>\n        <button id=\"close-button\">Close</button>\n\n        <script>\n          document.getElementById('close-button').onclick = function() {\n            Office.context.ui.messageParent('Dialog closed');\n          };\n        <\/script>\n      </body>\n    </html>\n  ");Office.context.ui.displayDialogAsync("data:text/html,".concat(encodeURIComponent(e)),{height:30,width:20},(function(t){t.status===Office.AsyncResultStatus.Failed&&console.error("Failed to display error dialog:",t.error.message)}))}