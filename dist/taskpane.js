/*! For license information please see taskpane.js.LICENSE.txt */
!function(){"use strict";var t={64796:function(t){t.exports="data:image/gif;base64,R0lGODlhAQABAIAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw=="},54058:function(t,e,r){t.exports=r.p+"1bcf162c09ad79190bfb.js"},58394:function(t,e,r){t.exports=r.p+"9367c0fb81875747ddc4.css"}},e={};function r(n){var o=e[n];if(void 0!==o)return o.exports;var i=e[n]={exports:{}};return t[n](i,i.exports,r),i.exports}r.m=t,r.d=function(t,e){for(var n in e)r.o(e,n)&&!r.o(t,n)&&Object.defineProperty(t,n,{enumerable:!0,get:e[n]})},r.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),r.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;r.g.importScripts&&(t=r.g.location+"");var e=r.g.document;if(!t&&e&&(e.currentScript&&"SCRIPT"===e.currentScript.tagName.toUpperCase()&&(t=e.currentScript.src),!t)){var n=e.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!t||!/^http(s?):/.test(t));)t=n[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),r.p=t}(),r.b=document.baseURI||self.location.href,function(){function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(){e=function(){return n};var r,n={},o=Object.prototype,i=o.hasOwnProperty,a=Object.defineProperty||function(t,e,r){t[e]=r.value},c="function"==typeof Symbol?Symbol:{},u=c.iterator||"@@iterator",s=c.asyncIterator||"@@asyncIterator",l=c.toStringTag||"@@toStringTag";function f(t,e,r){return Object.defineProperty(t,e,{value:r,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{f({},"")}catch(r){f=function(t,e,r){return t[e]=r}}function p(t,e,r,n){var o=e&&e.prototype instanceof b?e:b,i=Object.create(o.prototype),c=new T(n||[]);return a(i,"_invoke",{value:S(t,r,c)}),i}function h(t,e,r){try{return{type:"normal",arg:t.call(e,r)}}catch(t){return{type:"throw",arg:t}}}n.wrap=p;var d="suspendedStart",y="suspendedYield",g="executing",v="completed",m={};function b(){}function w(){}function x(){}var E={};f(E,u,(function(){return this}));var A=Object.getPrototypeOf,L=A&&A(A(R([])));L&&L!==o&&i.call(L,u)&&(E=L);var O=x.prototype=b.prototype=Object.create(E);function P(t){["next","throw","return"].forEach((function(e){f(t,e,(function(t){return this._invoke(e,t)}))}))}function j(e,r){function n(o,a,c,u){var s=h(e[o],e,a);if("throw"!==s.type){var l=s.arg,f=l.value;return f&&"object"==t(f)&&i.call(f,"__await")?r.resolve(f.__await).then((function(t){n("next",t,c,u)}),(function(t){n("throw",t,c,u)})):r.resolve(f).then((function(t){l.value=t,c(l)}),(function(t){return n("throw",t,c,u)}))}u(s.arg)}var o;a(this,"_invoke",{value:function(t,e){function i(){return new r((function(r,o){n(t,e,r,o)}))}return o=o?o.then(i,i):i()}})}function S(t,e,n){var o=d;return function(i,a){if(o===g)throw Error("Generator is already running");if(o===v){if("throw"===i)throw a;return{value:r,done:!0}}for(n.method=i,n.arg=a;;){var c=n.delegate;if(c){var u=k(c,n);if(u){if(u===m)continue;return u}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(o===d)throw o=v,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);o=g;var s=h(t,e,n);if("normal"===s.type){if(o=n.done?v:y,s.arg===m)continue;return{value:s.arg,done:n.done}}"throw"===s.type&&(o=v,n.method="throw",n.arg=s.arg)}}}function k(t,e){var n=e.method,o=t.iterator[n];if(o===r)return e.delegate=null,"throw"===n&&t.iterator.return&&(e.method="return",e.arg=r,k(t,e),"throw"===e.method)||"return"!==n&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+n+"' method")),m;var i=h(o,t.iterator,e.arg);if("throw"===i.type)return e.method="throw",e.arg=i.arg,e.delegate=null,m;var a=i.arg;return a?a.done?(e[t.resultName]=a.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=r),e.delegate=null,m):a:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,m)}function I(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function _(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function T(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(I,this),this.reset(!0)}function R(e){if(e||""===e){var n=e[u];if(n)return n.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,a=function t(){for(;++o<e.length;)if(i.call(e,o))return t.value=e[o],t.done=!1,t;return t.value=r,t.done=!0,t};return a.next=a}}throw new TypeError(t(e)+" is not iterable")}return w.prototype=x,a(O,"constructor",{value:x,configurable:!0}),a(x,"constructor",{value:w,configurable:!0}),w.displayName=f(x,l,"GeneratorFunction"),n.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===w||"GeneratorFunction"===(e.displayName||e.name))},n.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,x):(t.__proto__=x,f(t,l,"GeneratorFunction")),t.prototype=Object.create(O),t},n.awrap=function(t){return{__await:t}},P(j.prototype),f(j.prototype,s,(function(){return this})),n.AsyncIterator=j,n.async=function(t,e,r,o,i){void 0===i&&(i=Promise);var a=new j(p(t,e,r,o),i);return n.isGeneratorFunction(e)?a:a.next().then((function(t){return t.done?t.value:a.next()}))},P(O),f(O,l,"Generator"),f(O,u,(function(){return this})),f(O,"toString",(function(){return"[object Generator]"})),n.keys=function(t){var e=Object(t),r=[];for(var n in e)r.push(n);return r.reverse(),function t(){for(;r.length;){var n=r.pop();if(n in e)return t.value=n,t.done=!1,t}return t.done=!0,t}},n.values=R,T.prototype={constructor:T,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=r,this.done=!1,this.delegate=null,this.method="next",this.arg=r,this.tryEntries.forEach(_),!t)for(var e in this)"t"===e.charAt(0)&&i.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=r)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function n(n,o){return c.type="throw",c.arg=t,e.next=n,o&&(e.method="next",e.arg=r),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var a=this.tryEntries[o],c=a.completion;if("root"===a.tryLoc)return n("end");if(a.tryLoc<=this.prev){var u=i.call(a,"catchLoc"),s=i.call(a,"finallyLoc");if(u&&s){if(this.prev<a.catchLoc)return n(a.catchLoc,!0);if(this.prev<a.finallyLoc)return n(a.finallyLoc)}else if(u){if(this.prev<a.catchLoc)return n(a.catchLoc,!0)}else{if(!s)throw Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return n(a.finallyLoc)}}}},abrupt:function(t,e){for(var r=this.tryEntries.length-1;r>=0;--r){var n=this.tryEntries[r];if(n.tryLoc<=this.prev&&i.call(n,"finallyLoc")&&this.prev<n.finallyLoc){var o=n;break}}o&&("break"===t||"continue"===t)&&o.tryLoc<=e&&e<=o.finallyLoc&&(o=null);var a=o?o.completion:{};return a.type=t,a.arg=e,o?(this.method="next",this.next=o.finallyLoc,m):this.complete(a)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),m},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.finallyLoc===t)return this.complete(r.completion,r.afterLoc),_(r),m}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.tryLoc===t){var n=r.completion;if("throw"===n.type){var o=n.arg;_(r)}return o}}throw Error("illegal catch attempt")},delegateYield:function(t,e,n){return this.delegate={iterator:R(t),resultName:e,nextLoc:n},"next"===this.method&&(this.arg=r),m}},n}function r(t,e,r,n,o,i,a){try{var c=t[i](a),u=c.value}catch(t){return void r(t)}c.done?e(u):Promise.resolve(u).then(n,o)}function n(){return o.apply(this,arguments)}function o(){var t;return t=e().mark((function t(){var r,n,o,a,c,u;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:if(t.prev=0,(r=document.getElementById("script-input").value).trim()){t.next=5;break}return i("Please enter a Python script to execute."),t.abrupt("return");case 5:return console.log("Sending script to Python server."),(n=document.getElementById("loading-indicator"))&&(n.style.display="block"),t.next=10,fetch("http://127.0.0.1:8000/execute-script",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({script:r})});case 10:if(o=t.sent,console.log("Received response status:",o.status),o.ok){t.next=17;break}return t.next=15,o.json();case 15:throw a=t.sent,new Error("Server error: ".concat(a.detail));case 17:return t.next=19,o.json();case 19:c=t.sent,console.log("Received plot data:",c.plot),n&&(n.style.display="none"),e=c.plot,s=void 0,(s=document.getElementById("plot-image"))?(s.src="data:image/png;base64,".concat(e),console.log("Plot image source updated.")):(console.error("'plot-image' element not found in the DOM."),i("Failed to display the plot image.")),t.next=31;break;case 25:t.prev=25,t.t0=t.catch(0),console.error("Error executing script:",t.t0),(u=document.getElementById("loading-indicator"))&&(u.style.display="none"),i("Failed to execute the Python script.");case 31:case"end":return t.stop()}var e,s}),t,null,[[0,25]])})),o=function(){var e=this,n=arguments;return new Promise((function(o,i){var a=t.apply(e,n);function c(t){r(a,o,i,c,u,"next",t)}function u(t){r(a,o,i,c,u,"throw",t)}c(void 0)}))},o.apply(this,arguments)}function i(t){var e='\n    <!DOCTYPE html>\n    <html>\n      <head>\n        <title>Error</title>\n        <style>\n          body { font-family: Arial, sans-serif; padding: 20px; }\n          .error { color: red; }\n          button { margin-top: 20px; padding: 10px 20px; }\n        </style>\n      </head>\n      <body>\n        <h2 class="error">Error</h2>\n        <p>'.concat(t,"</p>\n        <button id=\"close-button\">Close</button>\n\n        <script>\n          document.getElementById('close-button').onclick = function() {\n            Office.context.ui.messageParent('Dialog closed');\n          };\n        <\/script>\n      </body>\n    </html>\n  ");Office.context.ui.displayDialogAsync("data:text/html,".concat(encodeURIComponent(e)),{height:30,width:20},(function(t){t.status===Office.AsyncResultStatus.Failed&&console.error("Failed to display error dialog:",t.error.message)}))}Office.onReady((function(t){if(console.log("Office.onReady called with info:",t),t.host===Office.HostType.Excel){console.log("Host is Excel. Proceeding to display app body."),document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex";var e=document.getElementById("execute-script");e?e.onclick=n:console.warn("'execute-script' button not found in the DOM.")}else console.log("Host is not Excel. Add-in not initialized.")}))}(),new URL(r(58394),r.b),new URL(r(54058),r.b),new URL(r(64796),r.b)}();
//# sourceMappingURL=taskpane.js.map