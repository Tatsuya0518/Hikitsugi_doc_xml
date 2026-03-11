# Third-party vendor libraries

## CodeMirror

- **Version**: 5.65.16
- **License**: MIT
- **Source**: https://www.npmjs.com/package/codemirror/v/5.65.16
- **Files**:
  - `codemirror.js` – core library
  - `codemirror.css` – core stylesheet
  - `xml.js` – XML language mode
  - `material.css` – Material theme stylesheet

To update, run:
```bash
npm install codemirror@<version>
cp node_modules/codemirror/lib/codemirror.js  static/vendor/codemirror/
cp node_modules/codemirror/lib/codemirror.css static/vendor/codemirror/
cp node_modules/codemirror/mode/xml/xml.js    static/vendor/codemirror/
cp node_modules/codemirror/theme/material.css static/vendor/codemirror/
```
