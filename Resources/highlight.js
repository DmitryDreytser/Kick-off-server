/*
Syntax highlighting with language autodetection.
https://highlightjs.org/
*/

(function(factory) {

  // Setup highlight.js for different environments. First is Node.js or
  // CommonJS.
  if(typeof exports !== 'undefined') {
    factory(exports);
  } else {
    // Export hljs globally even when using AMD for cases when this script
    // is loaded with others that may still expect a global hljs.
    window.hljs = factory({});

    // Finally register the global hljs with AMD.
    if(typeof define === 'function' && define.amd) {
      define([], function() {
        return window.hljs;
      });
    }
  }

}(function(hljs) {

    if (!String.prototype.trim) {
        (function () {
            // Make sure we trim BOM and NBSP
            var rtrim = /^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g;
            String.prototype.trim = function () {
                return this.replace(rtrim, '');
            };
        })();
    }

    ////////////////////////////////////////////

    function createObj(base, props) {
        var inst;
        if (Object.create) {
            inst = Object.create(base, props);
        } else {
            var ctor = function () { };
            ctor.prototype = base;
            inst = new ctor();
            if (props) copyObj(props, inst);
        }
        
        return inst;
    };

    function copyObj(obj, target, overwrite) {
        if (!target) target = {};
        for (var prop in obj)
            if (obj.hasOwnProperty(prop) && (overwrite !== false || !target.hasOwnProperty(prop)))
                target[prop] = obj[prop];
        return target;
    }

    if (!Array.prototype.indexOf) {
        Array.prototype.indexOf = function (searchElement /*, fromIndex */) {
            "use strict";
            if (this == null) {
                throw new TypeError();
            }
            var t = Object(this);
            var len = t.length >>> 0;
            if (len === 0) {
                return -1;
            }
            var n = 0;
            if (arguments.length > 1) {
                n = Number(arguments[1]);
                if (n != n) { // shortcut for verifying if it's NaN
                    n = 0;
                } else if (n != 0 && n != Infinity && n != -Infinity) {
                    n = (n > 0 || -1) * Math.floor(Math.abs(n));
                }
            }
            if (n >= len) {
                return -1;
            }
            var k = n >= 0 ? n : Math.max(len - Math.abs(n), 0);
            for (; k < len; k++) {
                if (k in t && t[k] === searchElement) {
                    return k;
                }
            }
            return -1;
        }
    }
    if (!Object.create && !Object.defineProperty) {
        //initIeObject();
    }


    function initIeObject() {
        var Constructor = XDomainRequest,
			defineProperty = Object.defineProperty,
			getOwnPropertyDescriptor = Object.getOwnPropertyDescriptor;

        Constructor.prototype.hasOwnProperty = function (key) {
            return !!~Object.getOwnPropertyNames(this).indexOf(key);
        };
        Object.defineProperties = function (object, descriptors) {
            for (var key in descriptors) {
                Object.defineProperty(object, key, descriptors[key]);
            }
        };
        JSON._stringify = JSON.stringify;
        JSON.stringify = function (object) {
            var props = object._settings_ && object._settings_.enumerable,
				o;
            if (props) {
                o = {};
                for (var i = 0; i < props.length; i++) {
                    o[props[i]] = object[props[i]];
                }
            } else {
                o = object;
            }
            return JSON._stringify(o);
        };

        if (!Object.keys) {
            Object.keys = (function () {
                return _.keys(this);
            });
        }

        Object.defineProperty = function (object, key, descriptor) {
            if (!(object instanceof Constructor)) return defineProperty(object, key, descriptor); // if another dom object

            var isConfugurable = descriptor.configurable !== undefined ? descriptor.configurable : false,
				_settings_ = object._settings_ = object._settings_ || {
				    enumerable: [],
				    notEnumerable: [],
				    notConfigurable: [],
				    descriptors: {}
				};

            delete descriptor.configurable;

            _settings_.descriptors[key] = descriptor;

            if (~_settings_.notConfigurable.indexOf(key)) {
                throw TypeError('Cannot redefine property: ' + key);
            }

            if (descriptor.enumerable === false) {
                if ('value' in descriptor) {
                    object['___' + key] = descriptor.value;
                    defineProperty(object, key, {
                        get: function () {
                            return this['___' + key];
                        },
                        set: function (value) {
                            this['___' + key] = value;
                        }
                    });
                    if (descriptor.writable === false) {
                        defineProperty(object, key, {
                            get: function () {
                                return descriptor.value;
                            }
                        });
                    }
                } else {
                    defineProperty(object, key, descriptor);
                }

                _settings_.notEnumerable.push(key);

            } else {
                if ('get' in descriptor || 'set' in descriptor) {
                    // что здесь можно сделать, чтоб свойство перечислялось с помощью for..in?
                    descriptor.enumerable === false;
                }

                defineProperty(object, key, descriptor);

                _settings_.enumerable.push(key);
            }

            if (isConfugurable === false) {
                _settings_.notConfigurable.push(key);
            }

            _settings_.descriptors[key].configurable = isConfugurable;
        };

        Object.getOwnPropertyNames = function (object) {
            return object._settings_.enumerable.concat(object._settings_.notEnumerable);
        };

        if (!Object.keys) {
            Object.keys = function (object) {
                return object._settings_.enumerable;
            };
        };

        Object.create = function (prototype, descriptors) {
            var object = new Constructor;

            prototype && extend(object, prototype);
            Object.defineProperties(object, descriptors);

            return object;
        };

        Object.getOwnPropertyDescriptor = function (object, key) {
            if (!(object instanceof Constructor)) return getOwnPropertyDescriptor(object, key); // if another dom object
            return object._settings_ && object._settings_.descriptors[key];
        };
    };

    function extend(o1, o2) {
        for (var i in o2) {
            o1[i] = o2[i];
        }
    }

    if (Object.defineProperty) {
        Object.each = function (object, callback) {
            var props = object._settings_ && object._settings_.enumerable;
            if (props) {
                for (var i = 0; i < props.length; i++) {
                    callback.call(object[props[i]], props[i], object[props[i]]);
                }
            } else {
                for (var i in object) if (object.hasOwnProperty(i)) {
                    callback.call(object[i], i, object[i]);
                }
            }
        };
    }

    ////////////////////////////////////////////


    if (!Object.keys) {
        Object.keys = (function () {
            return _.keys(this);
        });
    }

    if (typeof Object.create != 'function') {
        //Object.create = (function() {
        //    var Object = function() {};
        //    return function (prototype) {
        //        if (arguments.length > 1) {
        //        //    throw Error('Second argument not supported');
        //        }
        //        if (typeof prototype != 'object') {
        //            throw TypeError('Argument must be an object');
        //        }
        //        Object.prototype = prototype;
        //        var result = new Object();
        //        Object.prototype = null;
        //        return result;
        //    };
        //})();
    }

    if (!Array.prototype.indexOf) {
        Array.prototype.indexOf = function (searchElement, fromIndex) {
            return _.indexOf(this, searchElement, fromIndex)
        };
    }

    if (!Array.prototype.filter) {
        Array.prototype.filter = function (fun, context) {
            return _.filter(this, fun, context);
        };
    }

    Array.prototype.diff = function (a) {
        return this.filter(function (i) {
            return !(a.indexOf(i) > -1);
        });
    };

    if (!Array.prototype.map) {

        Array.prototype.map = function (callback, thisArg) {

            return _.map(this, callback, thisArg);
        };
    }

    if (!Array.prototype.forEach) {
        Array.prototype.forEach = function (callback, thisArg) {
            _.forEach(this, callback, thisArg);
        };
    }


  /* Utility functions */

  function escape(value) {
    return value.replace(/&/gm, '&amp;').replace(/</gm, '&lt;').replace(/>/gm, '&gt;');
  }

  function tag(node) {
    return node.nodeName.toLowerCase();
  }

  function testRe(re, lexeme) {
    var match = re && re.exec(lexeme);
    return match && match.index == 0;
  }

  function blockLanguage(block) {
    var classes = (block.className + ' ' + (block.parentNode ? block.parentNode.className : '')).split(/\s+/);
    classes = classes.map(function(c) {return c.replace(/^lang(uage)?-/, '');});
    return classes.filter(function(c) {return getLanguage(c) || /no(-?)highlight/.test(c);})[0];
  }

  function inherit(parent, obj) {
    var result = {};
    for (var key in parent)
      result[key] = parent[key];
    if (obj)
      for (var key in obj)
        result[key] = obj[key];
    return result;
  };

  /* Stream merging */

  function nodeStream(node) {
    var result = [];
    (function _nodeStream(node, offset) {
      for (var child = node.firstChild; child; child = child.nextSibling) {
        if (child.nodeType == 3)
          offset += child.nodeValue.length;
        else if (child.nodeType == 1) {
          result.push({
            event: 'start',
            offset: offset,
            node: child
          });
          offset = _nodeStream(child, offset);
          // Prevent void elements from having an end tag that would actually
          // double them in the output. There are more void elements in HTML
          // but we list only those realistically expected in code display.
          if (!tag(child).match(/br|hr|img|input/)) {
            result.push({
              event: 'stop',
              offset: offset,
              node: child
            });
          }
        }
      }
      return offset;
    })(node, 0);
    return result;
  }

  function mergeStreams(original, highlighted, value) {
    var processed = 0;
    var result = '';
    var nodeStack = [];

    function selectStream() {
      if (!original.length || !highlighted.length) {
        return original.length ? original : highlighted;
      }
      if (original[0].offset != highlighted[0].offset) {
        return (original[0].offset < highlighted[0].offset) ? original : highlighted;
      }

      /*
      To avoid starting the stream just before it should stop the order is
      ensured that original always starts first and closes last:

      if (event1 == 'start' && event2 == 'start')
        return original;
      if (event1 == 'start' && event2 == 'stop')
        return highlighted;
      if (event1 == 'stop' && event2 == 'start')
        return original;
      if (event1 == 'stop' && event2 == 'stop')
        return highlighted;

      ... which is collapsed to:
      */
      return highlighted[0].event == 'start' ? original : highlighted;
    }

    function open(node) {
      function attr_str(a) {return ' ' + a.nodeName + '="' + escape(a.value) + '"';}
      result += '<' + tag(node) + Array.prototype.map.call(node.attributes, attr_str).join('') + '>';
    }

    function close(node) {
      result += '</' + tag(node) + '>';
    }

    function render(event) {
      (event.event == 'start' ? open : close)(event.node);
    }

    while (original.length || highlighted.length) {
      var stream = selectStream();
      result += escape(value.substr(processed, stream[0].offset - processed));
      processed = stream[0].offset;
      if (stream == original) {
        /*
        On any opening or closing tag of the original markup we first close
        the entire highlighted node stack, then render the original tag along
        with all the following original tags at the same offset and then
        reopen all the tags on the highlighted stack.
        */
        nodeStack.reverse().forEach(close);
        do {
          render(stream.splice(0, 1)[0]);
          stream = selectStream();
        } while (stream == original && stream.length && stream[0].offset == processed);
        nodeStack.reverse().forEach(open);
      } else {
        if (stream[0].event == 'start') {
          nodeStack.push(stream[0].node);
        } else {
          nodeStack.pop();
        }
        render(stream.splice(0, 1)[0]);
      }
    }
    return result + escape(value.substr(processed));
  }

  /* Initialization */

  function compileLanguage(language) {

    function reStr(re) {
        return (re && re.source) || re;
    }

    function langRe(value, global) {
      return RegExp(
        reStr(value),
        'm' + (language.case_insensitive ? 'i' : '') + (global ? 'g' : '')
      );
    }

    function compileMode(mode, parent) {
      if (mode.compiled)
        return;
      mode.compiled = true;

      mode.keywords = mode.keywords || mode.beginKeywords;
      if (mode.keywords) {
        var compiled_keywords = {};

        var flatten = function (className, str)
        {
            if (str)
            {
                if (language.case_insensitive)
                {
                    str = str.toLowerCase();
                }

                str.split(' ').forEach(function (kw)
                {
                    var pair = kw.split('|');
                    compiled_keywords[pair[0]] = [className, pair[1] ? Number(pair[1]) : 1];
                });
            }
        };

        if (typeof mode.keywords == 'string') { // string
          flatten('keyword', mode.keywords);
        } else {
          Object.keys(mode.keywords).forEach(function (className) {
            flatten(className, mode.keywords[className]);
          });
        }
        mode.keywords = compiled_keywords;
      }
      mode.lexemesRe = langRe(mode.lexemes || /\b\w+\b/, true);

      if (parent) {
        if (mode.beginKeywords) {
          mode.begin = '\\b(' + mode.beginKeywords.split(' ').join('|') + ')\\b';
        }
        if (!mode.begin)
          mode.begin = /\B|\b/;
        mode.beginRe = langRe(mode.begin);
        if (!mode.end && !mode.endsWithParent)
          mode.end = /\B|\b/;
        if (mode.end)
          mode.endRe = langRe(mode.end);
        mode.terminator_end = reStr(mode.end) || '';
        if (mode.endsWithParent && parent.terminator_end)
          mode.terminator_end += (mode.end ? '|' : '') + parent.terminator_end;
      }
      if (mode.illegal)
        mode.illegalRe = langRe(mode.illegal);
      if (mode.relevance === undefined)
        mode.relevance = 1;
      if (!mode.contains) {
        mode.contains = [];
      }
      var expanded_contains = [];
      mode.contains.forEach(function(c) {
        if (c.variants) {
          c.variants.forEach(function(v) {expanded_contains.push(inherit(c, v));});
        } else {
          expanded_contains.push(c == 'self' ? mode : c);
        }
      });
      mode.contains = expanded_contains;
      mode.contains.forEach(function(c) {compileMode(c, mode);});

      if (mode.starts) {
        compileMode(mode.starts, parent);
      }

      var terminators =
        mode.contains.map(function(c) {
          return c.beginKeywords ? '\\.?(' + c.begin + ')\\.?' : c.begin;
        })
        .concat([mode.terminator_end, mode.illegal])
        .map(reStr)
        .filter(Boolean);
      mode.terminators = terminators.length ? langRe(terminators.join('|'), true) : {exec: function(s) {return null;}};
    }

    compileMode(language);
  }

  /*
  Core highlighting function. Accepts a language name, or an alias, and a
  string with the code to highlight. Returns an object with the following
  properties:

  - relevance (int)
  - value (an HTML string with highlighting markup)

  */
  function highlight(name, value, ignore_illegals, continuation) {

    function subMode(lexeme, mode) {
      for (var i = 0; i < mode.contains.length; i++) {
        if (testRe(mode.contains[i].beginRe, lexeme)) {
          return mode.contains[i];
        }
      }
    }

    function endOfMode(mode, lexeme) {
      if (testRe(mode.endRe, lexeme)) {
        return mode;
      }
      if (mode.endsWithParent) {
        return endOfMode(mode.parent, lexeme);
      }
    }

    function isIllegal(lexeme, mode) {
      return !ignore_illegals && testRe(mode.illegalRe, lexeme);
    }

    function keywordMatch(mode, match) {
      var match_str = language.case_insensitive ? match[0].toLowerCase() : match[0];
      return mode.keywords.hasOwnProperty(match_str) && mode.keywords[match_str];
    }

    function buildSpan(classname, insideSpan, leaveOpen, noPrefix) {
      var classPrefix = noPrefix ? '' : options.classPrefix,
          openSpan    = '<span class="' + classPrefix,
          closeSpan   = leaveOpen ? '' : '</span>';

      openSpan += classname + '">';

      return openSpan + insideSpan + closeSpan;
    }

    function processKeywords() {
      if (!top.keywords)
        return escape(mode_buffer);
      var result = '';
      var last_index = 0;
      top.lexemesRe.lastIndex = 0;
      var match = top.lexemesRe.exec(mode_buffer);
      while (match) {
        result += escape(mode_buffer.substr(last_index, match.index - last_index));
        var keyword_match = keywordMatch(top, match);
        if (keyword_match) {
          relevance += keyword_match[1];
          result += buildSpan(keyword_match[0], escape(match[0]));
        } else {
          result += escape(match[0]);
        }
        last_index = top.lexemesRe.lastIndex;
        match = top.lexemesRe.exec(mode_buffer);
      }
      return result + escape(mode_buffer.substr(last_index));
    }

    function processSubLanguage() {
      if (top.subLanguage && !languages[top.subLanguage]) {
        return escape(mode_buffer);
      }
      var result = top.subLanguage ? highlight(top.subLanguage, mode_buffer, true, continuations[top.subLanguage]) : highlightAuto(mode_buffer);
      // Counting embedded language score towards the host language may be disabled
      // with zeroing the containing mode relevance. Usecase in point is Markdown that
      // allows XML everywhere and makes every XML snippet to have a much larger Markdown
      // score.
      if (top.relevance > 0) {
        relevance += result.relevance;
      }
      if (top.subLanguageMode == 'continuous') {
        continuations[top.subLanguage] = result.top;
      }
      return buildSpan(result.language, result.value, false, true);
    }

    function processBuffer() {
      return top.subLanguage !== undefined ? processSubLanguage() : processKeywords();
    }

    function startNewMode(mode, lexeme) {
      var markup = mode.className? buildSpan(mode.className, '', true): '';
      if (mode.returnBegin) {
        result += markup;
        mode_buffer = '';
      } else if (mode.excludeBegin) {
        result += escape(lexeme) + markup;
        mode_buffer = '';
      } else {
        result += markup;
        mode_buffer = lexeme;
      }
      top = createObj(mode, { parent: { value: top } }); // Object.Create
    }

    function processLexeme(buffer, lexeme) {

      mode_buffer += buffer;
      if (lexeme === undefined) {
        result += processBuffer();
        return 0;
      }

      var new_mode = subMode(lexeme, top);
      if (new_mode) {
        result += processBuffer();
        startNewMode(new_mode, lexeme);
        return new_mode.returnBegin ? 0 : lexeme.length;
      }

      var end_mode = endOfMode(top, lexeme);
      if (end_mode) {
        var origin = top;
        if (!(origin.returnEnd || origin.excludeEnd)) {
          mode_buffer += lexeme;
        }
        result += processBuffer();
        do {
          if (top.className) {
            result += '</span>';
          }
          relevance += top.relevance;
          top = top.parent;
        } while (top != end_mode.parent);
        if (origin.excludeEnd) {
          result += escape(lexeme);
        }
        mode_buffer = '';
        if (end_mode.starts) {
          startNewMode(end_mode.starts, '');
        }
        return origin.returnEnd ? 0 : lexeme.length;
      }

      if (isIllegal(lexeme, top))
        throw new Error('Illegal lexeme "' + lexeme + '" for mode "' + (top.className || '<unnamed>') + '"');

      /*
      Parser should not reach this point as all types of lexemes should be caught
      earlier, but if it does due to some bug make sure it advances at least one
      character forward to prevent infinite looping.
      */
      mode_buffer += lexeme;
      return lexeme.length || 1;
    }

    var language = getLanguage(name);
    if (!language) {
      throw new Error('Unknown language: "' + name + '"');
    }

    compileLanguage(language);
    var top = continuation || language;
    var continuations = {}; // keep continuations for sub-languages
    var result = '';
    for(var current = top; current != language; current = current.parent) {
      if (current.className) {
        result = buildSpan(current.className, '', true) + result;
      }
    }
    var mode_buffer = '';
    var relevance = 0;
    try {
      var match, count, index = 0;
      while (top.terminators) {
        top.terminators.lastIndex = index;
        match = top.terminators.exec(value);
        if (!match)
          break;
        count = processLexeme(value.substr(index, match.index - index), match[0]);
        index = match.index + count;
      }
      processLexeme(value.substr(index));
      for(var current = top; current.parent; current = current.parent) { // close dangling modes
        if (current.className) {
          result += '</span>';
        }
      };
      return {
        relevance: relevance,
        value: result,
        language: name,
        top: top
      };
    } catch (e) {
        if (e.message.indexOf('Illegal') != -1)
        {
            return {relevance: 0, value: escape(value) };
        }
          else
          {
            throw e;
          }
    }
  }

  /*
  Highlighting with language detection. Accepts a string with the code to
  highlight. Returns an object with the following properties:

  - language (detected language)
  - relevance (int)
  - value (an HTML string with highlighting markup)
  - second_best (object with the same structure for second-best heuristically
    detected language, may be absent)

  */
  function highlightAuto(text, languageSubset) {
    languageSubset = languageSubset || options.languages || Object.keys(languages);
    var result = {
      relevance: 0,
      value: escape(text)
    };
    var second_best = result;
    languageSubset.forEach(function(name) {
      if (!getLanguage(name)) {
        return;
      }
      var current = highlight(name, text, false);
      current.language = name;
      if (current.relevance > second_best.relevance) {
        second_best = current;
      }
      if (current.relevance > result.relevance) {
        second_best = result;
        result = current;
      }
    });
    if (second_best.language) {
      result.second_best = second_best;
    }
    return result;
  }

  /*
  Post-processing of the highlighted markup:

  - replace TABs with something more useful
  - replace real line-breaks with '<br>' for non-pre containers

  */
  function fixMarkup(value) {
    if (options.tabReplace) {
      value = value.replace(/^((<[^>]+>|\t)+)/gm, function(match, p1, offset, s) {
        return p1.replace(/\t/g, options.tabReplace);
      });
    }
    if (options.useBR) {
      value = value.replace(/\n/g, '<br>');
    }
    return value;
  }

  function buildClassName(prevClassName, currentLang, resultLang) {
    var language = currentLang ? aliases[currentLang] : resultLang,
        result   = [prevClassName.trim()];

    if (!prevClassName.match(/\bhljs\b/)) {
      result.push('hljs');
    }

    if (language) {
      result.push(language);
    }

    return result.join(' ').trim();
  }

  /*
  Applies highlighting to a DOM node containing code. Accepts a DOM node and
  two optional parameters for fixMarkup.
  */
  function highlightBlock(block) {
    var language = blockLanguage(block);
    if (/no(-?)highlight/.test(language))
        return;

    var node;
    if (options.useBR) {
      node = document.createElementNS('http://www.w3.org/1999/xhtml', 'div');
      node.innerHTML = block.innerHTML.replace(/\n/g, '').replace(/<br[ \/]*>/g, '\n');
    } else {
      node = block;
    }

    var text = node.textContent ? node.textContent : node.innerHTML;
    if (!text)
        return;

    var result = language ? highlight(language, text, true) : highlightAuto(text);

    var originalStream = nodeStream(node);
    if (originalStream.length) {
      var resultNode = document.createElementNS('http://www.w3.org/1999/xhtml', 'div');
      resultNode.innerHTML = result.value;
      result.value = mergeStreams(originalStream, nodeStream(resultNode), text);
    }
    result.value = fixMarkup(result.value);

    block.innerHTML = result.value;
    block.className = buildClassName(block.className, language, result.language);
    block.result = {
      language: result.language,
      re: result.relevance
    };
    if (result.second_best) {
      block.second_best = {
        language: result.second_best.language,
        re: result.second_best.relevance
      };
    }
  }

  var options = {
    classPrefix: 'hljs-',
    tabReplace: null,
    useBR: false,
    languages: undefined
  };

  /*
  Updates highlight.js global options with values passed in the form of an object
  */
  function configure(user_options) {
    options = inherit(options, user_options);
  }

  /*
  Applies highlighting to all <pre><code>..</code></pre> blocks on a page.
  */
  function initHighlighting() {
    if (initHighlighting.called)
      return;
    initHighlighting.called = true;

    var blocks = $('pre code');
    Array.prototype.forEach.call(blocks, highlightBlock);
  }

  /*
  Attaches highlighting to the page load event.
  */
  function initHighlightingOnLoad() {
    addEventListener('DOMContentLoaded', initHighlighting, false);
    addEventListener('load', initHighlighting, false);
  }

  var languages = {};
  var aliases = {};

  function registerLanguage(name, language) {
    var lang = languages[name] = language(hljs);
    if (lang.aliases) {
      lang.aliases.forEach(function(alias) {aliases[alias] = name;});
    }
  }

  function listLanguages() {
    return Object.keys(languages);
  }

  function getLanguage(name) {
    return languages[name] || languages[aliases[name]];
  }

  /* Interface definition */

  hljs.highlight = highlight;
  hljs.highlightAuto = highlightAuto;
  hljs.fixMarkup = fixMarkup;
  hljs.highlightBlock = highlightBlock;
  hljs.configure = configure;
  hljs.initHighlighting = initHighlighting;
  hljs.initHighlightingOnLoad = initHighlightingOnLoad;
  hljs.registerLanguage = registerLanguage;
  hljs.listLanguages = listLanguages;
  hljs.getLanguage = getLanguage;
  hljs.inherit = inherit;

  // Common regexps
  hljs.IDENT_RE = '[a-zA-Z]\\w*';
  hljs.UNDERSCORE_IDENT_RE = '[a-zA-Z_]\\w*';
  hljs.NUMBER_RE = '\\b\\d+(\\.\\d+)?';
  hljs.C_NUMBER_RE = '\\b(0[xX][a-fA-F0-9]+|(\\d+(\\.\\d*)?|\\.\\d+)([eE][-+]?\\d+)?)'; // 0x..., 0..., decimal, float
  hljs.BINARY_NUMBER_RE = '\\b(0b[01]+)'; // 0b...
  hljs.RE_STARTERS_RE = '!|!=|!==|%|%=|&|&&|&=|\\*|\\*=|\\+|\\+=|,|-|-=|/=|/|:|;|<<|<<=|<=|<|===|==|=|>>>=|>>=|>=|>>>|>>|>|\\[|\\{|\\(|\\^|\\^=|\\||\\|=|\\|\\||~';

  // Common modes
  hljs.BACKSLASH_ESCAPE = {
    begin: '\\\\[\\s\\S]', relevance: 0
  };
  hljs.APOS_STRING_MODE = {
    className: 'string',
    begin: '\'', end: '\'',
    illegal: '\\n',
    contains: [hljs.BACKSLASH_ESCAPE]
  };
  hljs.QUOTE_STRING_MODE = {
    className: 'string',
    begin: '"', end: '"',
    illegal: '\\n',
    contains: [hljs.BACKSLASH_ESCAPE]
  };
  hljs.PHRASAL_WORDS_MODE = {
    begin: /\b(a|an|the|are|I|I'm|isn't|don't|doesn't|won't|but|just|should|pretty|simply|enough|gonna|going|wtf|so|such)\b/
  };
  hljs.C_LINE_COMMENT_MODE = {
    className: 'comment',
    begin: '//', end: '$',
    contains: [hljs.PHRASAL_WORDS_MODE]
  };
  hljs.C_BLOCK_COMMENT_MODE = {
    className: 'comment',
    begin: '/\\*', end: '\\*/',
    contains: [hljs.PHRASAL_WORDS_MODE]
  };
  hljs.HASH_COMMENT_MODE = {
    className: 'comment',
    begin: '#', end: '$',
    contains: [hljs.PHRASAL_WORDS_MODE]
  };
  hljs.NUMBER_MODE = {
    className: 'number',
    begin: hljs.NUMBER_RE,
    relevance: 0
  };
  hljs.C_NUMBER_MODE = {
    className: 'number',
    begin: hljs.C_NUMBER_RE,
    relevance: 0
  };
  hljs.BINARY_NUMBER_MODE = {
    className: 'number',
    begin: hljs.BINARY_NUMBER_RE,
    relevance: 0
  };
  hljs.CSS_NUMBER_MODE = {
    className: 'number',
    begin: hljs.NUMBER_RE + '(' +
      '%|em|ex|ch|rem'  +
      '|vw|vh|vmin|vmax' +
      '|cm|mm|in|pt|pc|px' +
      '|deg|grad|rad|turn' +
      '|s|ms' +
      '|Hz|kHz' +
      '|dpi|dpcm|dppx' +
      ')?',
    relevance: 0
  };
  hljs.REGEXP_MODE = {
    className: 'regexp',
    begin: /\//, end: /\/[gimuy]*/,
    illegal: /\n/,
    contains: [
      hljs.BACKSLASH_ESCAPE,
      {
        begin: /\[/, end: /\]/,
        relevance: 0,
        contains: [hljs.BACKSLASH_ESCAPE]
      }
    ]
  };
  hljs.TITLE_MODE = {
    className: 'title',
    begin: hljs.IDENT_RE,
    relevance: 0
  };
  hljs.UNDERSCORE_TITLE_MODE = {
    className: 'title',
    begin: hljs.UNDERSCORE_IDENT_RE,
    relevance: 0
  };

  return hljs;
}));

hljs.registerLanguage('1c',  function(hljs){
    var IDENT_RE_RU = '[a-zA-Z\u0430-\u044f\u0410-\u042f][a-zA-Z0-9_\u0430-\u044f\u0410-\u042f]*';
    var OneS_KEYWORDS = '\u0432\u043e\u0437\u0432\u0440\u0430\u0442 \u0434\u0430\u0442\u0430 \u0434\u043b\u044f \u0435\u0441\u043b\u0438 \u0438 \u0438\u043b\u0438 \u0438\u043d\u0430\u0447\u0435 \u0438\u043d\u0430\u0447\u0435\u0435\u0441\u043b\u0438 \u0438\u0441\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435 \u043a\u043e\u043d\u0435\u0446\u0435\u0441\u043b\u0438 \u043a\u043e\u043d\u0435\u0446\u043f\u043e\u043f\u044b\u0442\u043a\u0438 \u043a\u043e\u043d\u0435\u0446\u043f\u0440\u043e\u0446\u0435\u0434\u0443\u0440\u044b \u043a\u043e\u043d\u0435\u0446\u0444\u0443\u043d\u043a\u0446\u0438\u0438 \u043a\u043e\u043d\u0435\u0446\u0446\u0438\u043a\u043b\u0430 \u043a\u043e\u043d\u0441\u0442\u0430\u043d\u0442\u0430 \u043d\u0435 \u043f\u0435\u0440\u0435\u0439\u0442\u0438 \u043f\u0435\u0440\u0435\u043c \u043f\u0435\u0440\u0435\u0447\u0438\u0441\u043b\u0435\u043d\u0438\u0435 \u043f\u043e \u043f\u043e\u043a\u0430 \u043f\u043e\u043f\u044b\u0442\u043a\u0430 \u043f\u0440\u0435\u0440\u0432\u0430\u0442\u044c \u043f\u0440\u043e\u0434\u043e\u043b\u0436\u0438\u0442\u044c \u043f\u0440\u043e\u0446\u0435\u0434\u0443\u0440\u0430 \u0441\u0442\u0440\u043e\u043a\u0430 \u0442\u043e\u0433\u0434\u0430 \u0444\u0441 \u0444\u0443\u043d\u043a\u0446\u0438\u044f \u0446\u0438\u043a\u043b \u0447\u0438\u0441\u043b\u043e \u044d\u043a\u0441\u043f\u043e\u0440\u0442';
    var OneS_BUILT_IN = 'ansitooem oemtoansi \u0432\u0432\u0435\u0441\u0442\u0438\u0432\u0438\u0434\u0441\u0443\u0431\u043a\u043e\u043d\u0442\u043e \u0432\u0432\u0435\u0441\u0442\u0438\u0434\u0430\u0442\u0443 \u0432\u0432\u0435\u0441\u0442\u0438\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435 \u0432\u0432\u0435\u0441\u0442\u0438\u043f\u0435\u0440\u0435\u0447\u0438\u0441\u043b\u0435\u043d\u0438\u0435 \u0432\u0432\u0435\u0441\u0442\u0438\u043f\u0435\u0440\u0438\u043e\u0434 \u0432\u0432\u0435\u0441\u0442\u0438\u043f\u043b\u0430\u043d\u0441\u0447\u0435\u0442\u043e\u0432 \u0432\u0432\u0435\u0441\u0442\u0438\u0441\u0442\u0440\u043e\u043a\u0443 \u0432\u0432\u0435\u0441\u0442\u0438\u0447\u0438\u0441\u043b\u043e \u0432\u043e\u043f\u0440\u043e\u0441 \u0432\u043e\u0441\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435 \u0432\u0440\u0435\u0433 \u0432\u044b\u0431\u0440\u0430\u043d\u043d\u044b\u0439\u043f\u043b\u0430\u043d\u0441\u0447\u0435\u0442\u043e\u0432 \u0432\u044b\u0437\u0432\u0430\u0442\u044c\u0438\u0441\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0435 \u0434\u0430\u0442\u0430\u0433\u043e\u0434 \u0434\u0430\u0442\u0430\u043c\u0435\u0441\u044f\u0446 \u0434\u0430\u0442\u0430\u0447\u0438\u0441\u043b\u043e \u0434\u043e\u0431\u0430\u0432\u0438\u0442\u044c\u043c\u0435\u0441\u044f\u0446 \u0437\u0430\u0432\u0435\u0440\u0448\u0438\u0442\u044c\u0440\u0430\u0431\u043e\u0442\u0443\u0441\u0438\u0441\u0442\u0435\u043c\u044b \u0437\u0430\u0433\u043e\u043b\u043e\u0432\u043e\u043a\u0441\u0438\u0441\u0442\u0435\u043c\u044b \u0437\u0430\u043f\u0438\u0441\u044c\u0436\u0443\u0440\u043d\u0430\u043b\u0430\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0430\u0446\u0438\u0438 \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u044c\u043f\u0440\u0438\u043b\u043e\u0436\u0435\u043d\u0438\u0435 \u0437\u0430\u0444\u0438\u043a\u0441\u0438\u0440\u043e\u0432\u0430\u0442\u044c\u0442\u0440\u0430\u043d\u0437\u0430\u043a\u0446\u0438\u044e \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435\u0432\u0441\u0442\u0440\u043e\u043a\u0443 \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435\u0432\u0441\u0442\u0440\u043e\u043a\u0443\u0432\u043d\u0443\u0442\u0440 \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435\u0432\u0444\u0430\u0439\u043b \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435\u0438\u0437\u0441\u0442\u0440\u043e\u043a\u0438 \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435\u0438\u0437\u0441\u0442\u0440\u043e\u043a\u0438\u0432\u043d\u0443\u0442\u0440 \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435\u0438\u0437\u0444\u0430\u0439\u043b\u0430 \u0438\u043c\u044f\u043a\u043e\u043c\u043f\u044c\u044e\u0442\u0435\u0440\u0430 \u0438\u043c\u044f\u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u0435\u043b\u044f \u043a\u0430\u0442\u0430\u043b\u043e\u0433\u0432\u0440\u0435\u043c\u0435\u043d\u043d\u044b\u0445\u0444\u0430\u0439\u043b\u043e\u0432 \u043a\u0430\u0442\u0430\u043b\u043e\u0433\u0438\u0431 \u043a\u0430\u0442\u0430\u043b\u043e\u0433\u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u0435\u043b\u044f \u043a\u0430\u0442\u0430\u043b\u043e\u0433\u043f\u0440\u043e\u0433\u0440\u0430\u043c\u043c\u044b \u043a\u043e\u0434\u0441\u0438\u043c\u0432 \u043a\u043e\u043c\u0430\u043d\u0434\u0430\u0441\u0438\u0441\u0442\u0435\u043c\u044b \u043a\u043e\u043d\u0433\u043e\u0434\u0430 \u043a\u043e\u043d\u0435\u0446\u043f\u0435\u0440\u0438\u043e\u0434\u0430\u0431\u0438 \u043a\u043e\u043d\u0435\u0446\u0440\u0430\u0441\u0441\u0447\u0438\u0442\u0430\u043d\u043d\u043e\u0433\u043e\u043f\u0435\u0440\u0438\u043e\u0434\u0430\u0431\u0438 \u043a\u043e\u043d\u0435\u0446\u0441\u0442\u0430\u043d\u0434\u0430\u0440\u0442\u043d\u043e\u0433\u043e\u0438\u043d\u0442\u0435\u0440\u0432\u0430\u043b\u0430 \u043a\u043e\u043d\u043a\u0432\u0430\u0440\u0442\u0430\u043b\u0430 \u043a\u043e\u043d\u043c\u0435\u0441\u044f\u0446\u0430 \u043a\u043e\u043d\u043d\u0435\u0434\u0435\u043b\u0438 \u043b\u0435\u0432 \u043b\u043e\u0433 \u043b\u043e\u043310 \u043c\u0430\u043a\u0441 \u043c\u0430\u043a\u0441\u0438\u043c\u0430\u043b\u044c\u043d\u043e\u0435\u043a\u043e\u043b\u0438\u0447\u0435\u0441\u0442\u0432\u043e\u0441\u0443\u0431\u043a\u043e\u043d\u0442\u043e \u043c\u0438\u043d \u043c\u043e\u043d\u043e\u043f\u043e\u043b\u044c\u043d\u044b\u0439\u0440\u0435\u0436\u0438\u043c \u043d\u0430\u0437\u0432\u0430\u043d\u0438\u0435\u0438\u043d\u0442\u0435\u0440\u0444\u0435\u0439\u0441\u0430 \u043d\u0430\u0437\u0432\u0430\u043d\u0438\u0435\u043d\u0430\u0431\u043e\u0440\u0430\u043f\u0440\u0430\u0432 \u043d\u0430\u0437\u043d\u0430\u0447\u0438\u0442\u044c\u0432\u0438\u0434 \u043d\u0430\u0437\u043d\u0430\u0447\u0438\u0442\u044c\u0441\u0447\u0435\u0442 \u043d\u0430\u0439\u0442\u0438 \u043d\u0430\u0439\u0442\u0438\u043f\u043e\u043c\u0435\u0447\u0435\u043d\u043d\u044b\u0435\u043d\u0430\u0443\u0434\u0430\u043b\u0435\u043d\u0438\u0435 \u043d\u0430\u0439\u0442\u0438\u0441\u0441\u044b\u043b\u043a\u0438 \u043d\u0430\u0447\u0430\u043b\u043e\u043f\u0435\u0440\u0438\u043e\u0434\u0430\u0431\u0438 \u043d\u0430\u0447\u0430\u043b\u043e\u0441\u0442\u0430\u043d\u0434\u0430\u0440\u0442\u043d\u043e\u0433\u043e\u0438\u043d\u0442\u0435\u0440\u0432\u0430\u043b\u0430 \u043d\u0430\u0447\u0430\u0442\u044c\u0442\u0440\u0430\u043d\u0437\u0430\u043a\u0446\u0438\u044e \u043d\u0430\u0447\u0433\u043e\u0434\u0430 \u043d\u0430\u0447\u043a\u0432\u0430\u0440\u0442\u0430\u043b\u0430 \u043d\u0430\u0447\u043c\u0435\u0441\u044f\u0446\u0430 \u043d\u0430\u0447\u043d\u0435\u0434\u0435\u043b\u0438 \u043d\u043e\u043c\u0435\u0440\u0434\u043d\u044f\u0433\u043e\u0434\u0430 \u043d\u043e\u043c\u0435\u0440\u0434\u043d\u044f\u043d\u0435\u0434\u0435\u043b\u0438 \u043d\u043e\u043c\u0435\u0440\u043d\u0435\u0434\u0435\u043b\u0438\u0433\u043e\u0434\u0430 \u043d\u0440\u0435\u0433 \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0430\u043e\u0436\u0438\u0434\u0430\u043d\u0438\u044f \u043e\u043a\u0440 \u043e\u043f\u0438\u0441\u0430\u043d\u0438\u0435\u043e\u0448\u0438\u0431\u043a\u0438 \u043e\u0441\u043d\u043e\u0432\u043d\u043e\u0439\u0436\u0443\u0440\u043d\u0430\u043b\u0440\u0430\u0441\u0447\u0435\u0442\u043e\u0432 \u043e\u0441\u043d\u043e\u0432\u043d\u043e\u0439\u043f\u043b\u0430\u043d\u0441\u0447\u0435\u0442\u043e\u0432 \u043e\u0441\u043d\u043e\u0432\u043d\u043e\u0439\u044f\u0437\u044b\u043a \u043e\u0442\u043a\u0440\u044b\u0442\u044c\u0444\u043e\u0440\u043c\u0443 \u043e\u0442\u043a\u0440\u044b\u0442\u044c\u0444\u043e\u0440\u043c\u0443\u043c\u043e\u0434\u0430\u043b\u044c\u043d\u043e \u043e\u0442\u043c\u0435\u043d\u0438\u0442\u044c\u0442\u0440\u0430\u043d\u0437\u0430\u043a\u0446\u0438\u044e \u043e\u0447\u0438\u0441\u0442\u0438\u0442\u044c\u043e\u043a\u043d\u043e\u0441\u043e\u043e\u0431\u0449\u0435\u043d\u0438\u0439 \u043f\u0435\u0440\u0438\u043e\u0434\u0441\u0442\u0440 \u043f\u043e\u043b\u043d\u043e\u0435\u0438\u043c\u044f\u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u0435\u043b\u044f \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c\u0432\u0440\u0435\u043c\u044f\u0442\u0430 \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c\u0434\u0430\u0442\u0443\u0442\u0430 \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c\u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0442\u0430 \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u044f\u043e\u0442\u0431\u043e\u0440\u0430 \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c\u043f\u043e\u0437\u0438\u0446\u0438\u044e\u0442\u0430 \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c\u043f\u0443\u0441\u0442\u043e\u0435\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435 \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c\u0442\u0430 \u043f\u0440\u0430\u0432 \u043f\u0440\u0430\u0432\u043e\u0434\u043e\u0441\u0442\u0443\u043f\u0430 \u043f\u0440\u0435\u0434\u0443\u043f\u0440\u0435\u0436\u0434\u0435\u043d\u0438\u0435 \u043f\u0440\u0435\u0444\u0438\u043a\u0441\u0430\u0432\u0442\u043e\u043d\u0443\u043c\u0435\u0440\u0430\u0446\u0438\u0438 \u043f\u0443\u0441\u0442\u0430\u044f\u0441\u0442\u0440\u043e\u043a\u0430 \u043f\u0443\u0441\u0442\u043e\u0435\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435 \u0440\u0430\u0431\u043e\u0447\u0430\u044f\u0434\u0430\u0442\u0442\u044c\u043f\u0443\u0441\u0442\u043e\u0435\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435 \u0440\u0430\u0431\u043e\u0447\u0430\u044f\u0434\u0430\u0442\u0430 \u0440\u0430\u0437\u0434\u0435\u043b\u0438\u0442\u0435\u043b\u044c\u0441\u0442\u0440\u0430\u043d\u0438\u0446 \u0440\u0430\u0437\u0434\u0435\u043b\u0438\u0442\u0435\u043b\u044c\u0441\u0442\u0440\u043e\u043a \u0440\u0430\u0437\u043c \u0440\u0430\u0437\u043e\u0431\u0440\u0430\u0442\u044c\u043f\u043e\u0437\u0438\u0446\u0438\u044e\u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0430 \u0440\u0430\u0441\u0441\u0447\u0438\u0442\u0430\u0442\u044c\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u044b\u043d\u0430 \u0440\u0430\u0441\u0441\u0447\u0438\u0442\u0430\u0442\u044c\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u044b\u043f\u043e \u0441\u0438\u0433\u043d\u0430\u043b \u0441\u0438\u043c\u0432 \u0441\u0438\u043c\u0432\u043e\u043b\u0442\u0430\u0431\u0443\u043b\u044f\u0446\u0438\u0438 \u0441\u043e\u0437\u0434\u0430\u0442\u044c\u043e\u0431\u044a\u0435\u043a\u0442 \u0441\u043e\u043a\u0440\u043b \u0441\u043e\u043a\u0440\u043b\u043f \u0441\u043e\u043a\u0440\u043f \u0441\u043e\u043e\u0431\u0449\u0438\u0442\u044c \u0441\u043e\u0441\u0442\u043e\u044f\u043d\u0438\u0435 \u0441\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435 \u0441\u0440\u0435\u0434 \u0441\u0442\u0430\u0442\u0443\u0441\u0432\u043e\u0437\u0432\u0440\u0430\u0442\u0430 \u0441\u0442\u0440\u0434\u043b\u0438\u043d\u0430 \u0441\u0442\u0440\u0437\u0430\u043c\u0435\u043d\u0438\u0442\u044c \u0441\u0442\u0440\u043a\u043e\u043b\u0438\u0447\u0435\u0441\u0442\u0432\u043e\u0441\u0442\u0440\u043e\u043a \u0441\u0442\u0440\u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c\u0441\u0442\u0440\u043e\u043a\u0443  \u0441\u0442\u0440\u0447\u0438\u0441\u043b\u043e\u0432\u0445\u043e\u0436\u0434\u0435\u043d\u0438\u0439 \u0441\u0444\u043e\u0440\u043c\u0438\u0440\u043e\u0432\u0430\u0442\u044c\u043f\u043e\u0437\u0438\u0446\u0438\u044e\u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0430 \u0441\u0447\u0435\u0442\u043f\u043e\u043a\u043e\u0434\u0443 \u0442\u0435\u043a\u0443\u0449\u0430\u044f\u0434\u0430\u0442\u0430 \u0442\u0435\u043a\u0443\u0449\u0435\u0435\u0432\u0440\u0435\u043c\u044f \u0442\u0438\u043f\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u044f \u0442\u0438\u043f\u0437\u043d\u0430\u0447\u0435\u043d\u0438\u044f\u0441\u0442\u0440 \u0443\u0434\u0430\u043b\u0438\u0442\u044c\u043e\u0431\u044a\u0435\u043a\u0442\u044b \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c\u0442\u0430\u043d\u0430 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c\u0442\u0430\u043f\u043e \u0444\u0438\u043a\u0441\u0448\u0430\u0431\u043b\u043e\u043d \u0444\u043e\u0440\u043c\u0430\u0442 \u0446\u0435\u043b \u0448\u0430\u0431\u043b\u043e\u043d';
    var DQUOTE =  {className: 'dquote',  begin: '""'};
    var STR_START = {
        className: 'string',
        begin: '"', end: '"|$',
        contains: [DQUOTE]
    };
    var STR_CONT = {
        className: 'string',
        begin: '\\|', end: '"|$',
        contains: [DQUOTE]
    };
  
    var BRACKETS = {className: 'brackets',  begin: '\\(|\\)', relevance: 0};
    var MATHS = { className: 'maths', begin: '\\+|\\-|\\=|\\*|\\<|\\>|\\.|\\>=|\\<=|\\,|\\;|\\x3F|\\x2F|\\x25|\\x5C', relevance: 0 };
  
    return {
        case_insensitive: true,
        lexemes: IDENT_RE_RU,
        keywords: {keyword: OneS_KEYWORDS, built_in: OneS_BUILT_IN},
        contains: [
          hljs.C_LINE_COMMENT_MODE,
          hljs.NUMBER_MODE,
          STR_START, STR_CONT,
          BRACKETS,
          MATHS,
          {
		
              className: 'function',
              begin: '(\u043f\u0440\u043e\u0446\u0435\u0434\u0443\u0440\u0430|\u0444\u0443\u043d\u043a\u0446\u0438\u044f)', end: '$',
              lexemes: IDENT_RE_RU,
              keywords: '\u043f\u0440\u043e\u0446\u0435\u0434\u0443\u0440\u0430 \u0444\u0443\u043d\u043a\u0446\u0438\u044f',
              contains: [
                hljs.inherit(hljs.TITLE_MODE, {begin: IDENT_RE_RU}),
                {
                    className: 'tail',
                    endsWithParent: true,
                    contains: [
                        BRACKETS,
                        MATHS,
                      {
                          className: 'params',
                          begin: '\\(', 
                          end: '\\)',
                          lexemes: IDENT_RE_RU,
                          keywords: '\u0437\u043d\u0430\u0447',
                          contains: [STR_START, STR_CONT]
                      },
                      {
                          className: 'export',
                          begin: '\u044d\u043a\u0441\u043f\u043e\u0440\u0442', endsWithParent: true,
                          lexemes: IDENT_RE_RU,
                          keywords: '\u044d\u043a\u0441\u043f\u043e\u0440\u0442',
                          contains: [hljs.C_LINE_COMMENT_MODE]
                      }			  
                    ]
                },
                hljs.C_LINE_COMMENT_MODE
              ]
          },
          {className: 'preprocessor', begin: '#', end: '$'},
          {className: 'date', begin: '\'\\d{2}\\.\\d{2}\\.(\\d{2}|\\d{4})\''}
        ]
    };
})
