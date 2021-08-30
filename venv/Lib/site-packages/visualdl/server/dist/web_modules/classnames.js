import{c as a}from"./common/_commonjsHelpers-4f955397.js";var c=a(function(f){/*!
  Copyright (c) 2018 Jed Watson.
  Licensed under the MIT License (MIT), see
  http://jedwatson.github.io/classnames
*/(function(){var l={}.hasOwnProperty;function t(){for(var e=[],o=0;o<arguments.length;o++){var s=arguments[o];if(!!s){var n=typeof s;if(n==="string"||n==="number")e.push(s);else if(Array.isArray(s)){if(s.length){var i=t.apply(null,s);i&&e.push(i)}}else if(n==="object")if(s.toString===Object.prototype.toString)for(var r in s)l.call(s,r)&&s[r]&&e.push(r);else e.push(s.toString())}}return e.join(" ")}f.exports?(t.default=t,f.exports=t):window.classNames=t})()});export default c;
