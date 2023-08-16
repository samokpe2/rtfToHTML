var _ = require('underscore');

class RtfElement {
    Indent(level) {
        for (let i = 0; i < level * 2; i++) {
            console.log("&nbsp;");
        }
    }
}

class RtfGroup extends RtfElement {
    constructor() {
        super();
        this.parent = null;
        this.children = [];
    }

    GetType() {
        if (this.children.length === 0) return null;

        const child = this.children[0];

        if (child instanceof RtfControlWord) {
            return child.word;
        } else if (child instanceof RtfControlSymbol) {
            return child.symbol === '*' ? '*' : null;
        }

        return null;
    }

    IsDestination() {
        if (this.children.length === 0) return null;

        const child = this.children[0];

        if (!(child instanceof RtfControlSymbol)) return null;

        return child.symbol === '*';
    }

    dump(level = 0) {
        console.log("<div>");
        this.Indent(level);
        console.log("{");
        console.log("</div>");

        for (const child of this.children) {
            if (child instanceof RtfGroup) {
                if (child.GetType() === "fonttbl") continue;
                if (child.GetType() === "colortbl") continue;
                if (child.GetType() === "stylesheet") continue;
                if (child.GetType() === "info") continue;
                if (child.GetType().startsWith("pict")) continue;
                if (child.IsDestination()) continue;
            }
            child.dump(level + 2);
        }

        console.log("<div>");
        this.Indent(level);
        console.log("}");
        console.log("</div>");
    }
}

class RtfControlWord extends RtfElement {
    constructor(word, parameter) {
        super();
        this.word = word;
        this.parameter = parameter;
    }

    dump(level) {
        console.log("<div style='color:green'>");
        this.Indent(level);
        console.log(`WORD ${this.word} (${this.parameter})`);
        console.log("</div>");
    }
}

class RtfControlSymbol extends RtfElement {
    constructor(symbol, parameter = 0) {
        super();
        this.symbol = symbol;
        this.parameter = parameter;
    }

    dump(level) {
        console.log("<div style='color:blue'>");
        this.Indent(level);
        console.log(`SYMBOL ${this.symbol} (${this.parameter})`);
        console.log("</div>");
    }
}

class RtfText extends RtfElement {
    constructor(text) {
        super();
        this.text = text;
    }

    dump(level) {
        console.log("<div style='color:red'>");
        this.Indent(level);
        console.log(`TEXT ${this.text}`);
        console.log("</div>");
    }
}

class RtfReader {
    constructor() {
        this.root = null;
    }

    GetChar() {
        this.char = null;
        if (this.pos < this.rtf.length) {
            this.char = this.rtf[this.pos++];
        } else {
            this.err = "Tried to read past EOF, RTF is probably truncated";
        }
    }

    ParseStartGroup() {
        const group = new RtfGroup();
    
        if (this.group !== null) {
            group.parent = this.group;
        }
    
        if (this.root === null) {
            // First group of the RTF document
            this.group = group;
            this.root = group;
            this.uc = [1]; // Create uc stack and insert the first default value
        } else {
            this.uc.push(this.uc[this.uc.length - 1]);
            this.group.children.push(group);
            this.group = group;
        }
    }

    isLetter() {
        const charCode = this.char.charCodeAt(0);
        return (charCode >= 65 && charCode <= 90) || (charCode >= 97 && charCode <= 122);
    }
    
    isDigit() {
        const charCode = this.char.charCodeAt(0);
        return charCode >= 48 && charCode <= 57;
    }
    
    isEndOfLine() {
        if (this.char === "\r" || this.char === "\n") {
            // Checks for a Windows/Acron type EOL
            if (this.rtf[this.pos] === "\n" || this.rtf[this.pos] === "\r") {
                this.GetChar();
            }
            return true;
        }
        return false;
    }
    
    isSpaceDelimiter() {
        return this.char === " " || this.isEndOfLine();
    }
    
    ParseEndGroup() {
        // Retrieve state of document from stack.
        this.group = this.group.parent;
        // Retrieve last uc value from stack
        this.uc.pop();
    }


    ParseControlWord() {

        this.GetChar();
        let word = "";
        while (this.isLetter()) {
            word += this.char;
            this.GetChar();
        }
    
        // Read parameter (if any) consisting of digits.
        // Parameter may be negative.
        let parameter = null;
        let negative = false;
        if (this.char === '-') {
            this.GetChar();
            negative = true;
        }
        while (this.isDigit()) {
            if (parameter === null) parameter = 0;
            parameter = parameter * 10 + parseInt(this.char);
            this.GetChar();
        }
        // if no parameter, assume control word's default (usually 1)
        // if no default, then assign 0 to the parameter
        if (parameter === null) parameter = 1;
    
        // convert to a negative number when applicable
        if (negative) parameter = -parameter;
    
        // Update uc value
        if (word === "uc") {
            this.uc.pop();
            this.uc.push(parameter);
        }
    
        // Skip space delimiter
        if (!this.isSpaceDelimiter()) this.pos--;
    
        // If this is \u, then the parameter will be followed
        // by ${this.uc} characters.
        if (word === "u") {
            // Convert parameter to unsigned decimal unicode
            if (negative) parameter = 65536 + parameter;
    
            // Will ignore replacement characters this.uc times
            let uc = this.uc[this.uc.length - 1];
            while (uc > 0) {
                this.GetChar();
                // If the replacement character is encoded as
                // hexadecimal value \'hh then jump over it
                if (this.char === '\\' && (this.rtf[this.pos] === "'" || this.rtf[this.pos] === '"')) {
                    this.pos += 3;
                } else if (this.char === '{' || this.char === '}') {
                    // Break if it's an RTF scope delimiter
                    break;
                }
                uc--;
            }
        }
    
        const rtfword = new RtfControlWord();
        rtfword.word = word;
        rtfword.parameter = parameter;
        this.group.children.push(rtfword);
    }
    
    ParseControlSymbol() {
        this.GetChar();
        // Read symbol (one character only).
        const symbol = this.char;
    
        // Symbols ordinarily have no parameter. However,
        // if this is \', then it is followed by a 2-digit hex-code:
        let parameter = 0;
        // Treat EOL symbols as \par control word
        if (this.isEndOfLine()) {
            const rtfword = new RtfControlWord();
            rtfword.word = 'par';
            rtfword.parameter = parameter;
            this.group.children.push(rtfword);
            return;
        } else if (symbol === "'") {
            this.GetChar();
            parameter = parseInt(this.char, 16);
        }
    
        const rtfsymbol = new RtfControlSymbol();
        rtfsymbol.symbol = symbol;
        rtfsymbol.parameter = parameter;
        this.group.children.push(rtfsymbol);
    }
    
    ParseControl() {
        // Beginning of an RTF control word or control symbol.
        // Look ahead by one character to see if it starts with
        // a letter (control word) or another symbol (control symbol):
        this.GetChar();
        this.pos--;
        if (this.isLetter())
            this.ParseControlWord();
        else
            this.ParseControlSymbol();
    }

    ParseText() {
        let text = "";
        let terminate = false;
        
        do {
            // Ignore EOL characters
            if (this.char === "\r" || this.char === "\n") {
                this.GetChar();
                continue;
            }
    
            // Is this an escape?
            if (this.char === '\\') {
                // Perform lookahead to see if this
                // is really an escape sequence.
                this.GetChar();
                switch (this.char) {
                    case '\\':
                    case '{':
                    case '}':
                        break;
                    default:
                        // Not an escape. Roll back.
                        this.pos -= 2;
                        terminate = true;
                        break;
                }
            } else if (this.char === '{' || this.char === '}') {
                this.pos--;
                terminate = true;
            }
    
            if (!terminate) {
                // Save plain text
                text += this.char;
                this.GetChar();
            }
        } while (!terminate && this.pos < this.len);
    
        const rtftext = new RtfText();
        rtftext.text = text;
    
        // If group does not exist, then this is not a valid RTF file.
        // Throw an exception.
        if (this.group !== null) {
            this.group.children.push(rtftext);
        }
    }

    Parse(rtf) {
        try {
            this.rtf = rtf;
            this.pos = 0;
            this.len = this.rtf.length;
            this.group = null;
            this.root = null;
    
            while (this.pos < this.len) {
                // Read next character:
                this.GetChar();
    
                // Ignore \r and \n
                if (this.char === "\n" || this.char === "\r") continue;
    
                // What type of character is this?
                switch (this.char) {
                    case '{':
                        this.ParseStartGroup();
                        break;
                    case '}':
                        this.ParseEndGroup();
                        break;
                    case '\\':
                        this.ParseControl();
                        break;
                    default:
                        this.ParseText();
                        break;
                }
            }
            
            return true;
        } catch (error) {
            return false;
        }
    }
}
    
class RtfImage {
    constructor() {
        this.Reset();
    }

    Reset() {
        this.format = 'bmp';
        this.width = 0; // in xExt if wmetafile otherwise in px
        this.height = 0; // in yExt if wmetafile otherwise in px
        this.goalWidth = 0; // in twips
        this.goalHeight = 0; // in twips
        this.pcScaleX = 100; // 100%
        this.pcScaleY = 100; // 100%
        this.binarySize = null; // Number of bytes of the binary data
        this.ImageData = null; // Binary or Hexadecimal Data
    }

    PrintImage() {
        // <img src="data:image/{FORMAT};base64,{#BDATA}" />
        let output = `<img src="data:image/${this.format};base64,`;
        
        if (this.binarySize !== null) { // process binary data
            // Process binary data here
            // Add appropriate code
        } else { // process hexadecimal data
            // If necessary, handle image format-specific logic here
            
            // Example for handling hexadecimal data (base64 encoding):
            output += Buffer.from(this.ImageData, 'hex').toString('base64');
        }

        output += `" />`;

        return output;
    }
    
}

class RtfFont {
    constructor() {
        this.fontFamily;
        this.fontName;
        this.charset;
        this.codePage;
    }
}

class RtfState {
    static fonttbl = {};
    static colortbl = {};
    static highlight = {
        1: 'Black',
        2: 'Blue',
        3: 'Cyan',
        4 : 'Green',
        5 : 'Magenta',
        6 : 'Red',
        7 : 'Yellow',
        8 : 'Unused',
        9 :  'DarkBlue',
        10 : 'DarkCyan',
        11 : 'DarkGreen',
        12 : 'DarkMagenta',
        13 : 'DarkRed',
        14 : 'DarkYellow',
        15 : 'DarkGray',
        16 : 'LightGray'
    };

    constructor() {
        this.Reset();
    }

    Reset(defaultFont = null) {
        this.bold = false;
        this.italic = false;
        this.underline = false;
        this.strike = false;
        this.hidden = false;
        this.fontsize = 0;
        this.fontcolor = null;
        this.background = null;
        this.font = defaultFont || null;
    }

    PrintStyle() {
        let style = "";

        if (this.bold) style += "font-weight:bold;";
        if (this.italic) style += "font-style:italic;";
        if (this.underline) style += "text-decoration:underline;";
        // state->underline is a toggle switch variable, no need for state->end_underline variable
        //if (this.state.end_underline) style += "text-decoration:none;";
        if (this.strike) style += "text-decoration:line-through;";
        if (this.hidden) style += "display:none;";
        
        if (this.font) {
            const fontFamily = RtfState.fonttbl[this.font]?.fontfamily;
            if (fontFamily) style += `font-family:${fontFamily};`;
        }

        // Background color:
        if (this.background) {
            const backgroundColor = RtfState.colortbl[this.background];
            if (backgroundColor) style += `background-color:${backgroundColor};`;
        }
        // Highlight color:
        else if (this.hcolor) {
            const highlightColor = RtfState.highlight[this.hcolor];
            if (highlightColor) style += `background-color:${highlightColor};`;
        }

        return style;
    }

    isLike(state) {
        if (!(state instanceof RtfState))
            return false;
        
        if (this.bold !== state.bold)
            return false;
        if (this.italic !== state.italic)
            return false;
        if (this.underline !== state.underline)
            return false;
        if (this.strike !== state.strike)
            return false;
        if (this.hidden !== state.hidden)
            return false;
        if (this.fontsize !== state.fontsize)
            return false;

        if (this.fontcolor) {
            if (!state.fontcolor)
                return false;
            if (this.fontcolor !== state.fontcolor)
                return false;
        } else if (state.fontcolor)
            return false;

        if (this.background) {
            if (!state.background)
                return false;
            if (this.background !== state.background)
                return false;
        } else if (state.background)
            return false;

        if (this.hcolor) {
            if (!state.hcolor)
                return false;
            if (this.hcolor !== state.hcolor)
                return false;
        } else if (state.hcolor)
            return false;

        if (this.font) {
            if (!state.font)
                return false;
        } else if (state.font)
            return false;

        return true;
    }
}

class RtfHtml {
    constructor(encoding = 'HTML-ENTITIES') {
        this.output = '';
        this.encoding;
        this.defaultFont;
        this.sup;
        this.table = 0;
        this.emf;
    
        if (encoding !== 'HTML-ENTITIES') {
            // Check if mbstring extension is loaded
            if (!require('mbstring')) {
                console.warn("PHP mbstring extension not enabled, reverting back to HTML-ENTITIES");
                encoding = 'HTML-ENTITIES';
            } else if (!mb_list_encodings().includes(encoding)) {
                console.warn("Unrecognized Encoding, reverting back to HTML-ENTITIES");
                encoding = 'HTML-ENTITIES';
            }
        }
        this.encoding = encoding;
    }

    Format(root) {
        // Keep track of style modifications
        this.previousState = null;
        // and create a stack of states
        this.states = [];
        // Put an initial standard state onto the stack
        this.state = new RtfState();
        this.states.push(this.state);
        // Keep track of opened html tags
        this.openedTags = { 'span': false, 'p': false };
        // Create the first paragraph
        this.OpenTag('p');
        // Begin format
        this.ProcessGroup(root);
        // Remove the last opened <p> tag and return
        return this.output.slice(0, -3);
    }

    ExtractFontTable(fontTblGrp) {
        // {' \fonttbl (<fontinfo> | ('{' <fontinfo> '}'))+ '}
        // <fontnum><fontfamily><fcharset>?<fprq>?<panose>?
        // <nontaggedname>?<fontemb>?<codepage>? <fontname><fontaltname>? ';'
        let fonttbl = {};
        const c = fontTblGrp.length;
    
        for (let i = 1; i < c; i++) {
            let fname = '';
            let fN = null;
            for (const child of fontTblGrp[i].children) {
                if (child instanceof RtfControlWord) {
                    switch (child.word) {
                        case 'f':
                            fN = child.parameter;
                            fonttbl[fN] = new RtfFont();
                            break;
                        // Font family names
                        case 'froman':
                            fonttbl[fN].fontFamily = 'serif';
                            break;
                        case 'fswiss':
                            fonttbl[fN].fontFamily = 'sans-serif';
                            break;
                        case 'fmodern':
                            fonttbl[fN].fontFamily = 'monospace';
                            break;
                        case 'fscript':
                            fonttbl[fN].fontFamily = 'cursive';
                            break;
                        case 'fdecor':
                            fonttbl[fN].fontFamily = 'fantasy';
                            break;
                        // case 'fnil': break; // default font
                        // case 'ftech': break; // symbol
                        // case 'fbidi': break; // bidirectional font
                        case 'fcharset': // charset
                            fonttbl[fN].charset =
                                this.GetEncodingFromCharset(child.parameter);
                            break;
                        case 'cpg': // code page
                            fonttbl[fN].codepage =
                                this.GetEncodingFromCodepage(child.parameter);
                            break;
                        case 'fprq': // Font pitch
                            fonttbl[fN].fprq = child.parameter;
                            break;
                        default:
                            break;
                    }
                } else if (child instanceof RtfText) {
                    // Save font name
                    fname += child.text;
                }
            }
            // Remove end ; delimiter from font name
            fonttbl[fN].fontName = fname.slice(0, -1);
    
            // Save extracted Font
            RtfState.fonttbl = fonttbl;
        }
    }
    
    ExtractColorTable(colorTblGrp) {
        // {\colortbl;\red0\green0\blue0;}
        // Index 0 of the RTF color table  is the 'auto' color
        let colortbl = [];
        const c = colorTblGrp.length;
        let color = '';
    
        for (let i = 1; i < c; i++) { // Iterate through colors
            if (colorTblGrp[i] instanceof RtfControlWord) {
                // Extract RGB color and convert it to hex string
                color = `#${colorTblGrp[i].parameter.toString(16).padStart(2, '0')}` +
                        `${colorTblGrp[i+1].parameter.toString(16).padStart(2, '0')}` +
                        `${colorTblGrp[i+2].parameter.toString(16).padStart(2, '0')}`;
                i += 2;
            } else if (colorTblGrp[i] instanceof RtfText) {
                // This is a delimiter ';', so
                if (i !== 1) { // Store the already extracted color
                    colortbl.push(color);
                } else { // This is the 'auto' color
                    colortbl.push(0);
                }
            }
        }
        RtfState.colortbl = colortbl;
    }
    
    ExtractImage(pictGrp) {
        const Image = new RtfImage();
        for (const child of pictGrp) {
            if (child instanceof RtfControlWord) {
                switch (child.word) {
                    // Picture Format
                    case "emfblip":
                        Image.format = 'emf';
                        this.emf = 1;
                        break;
                    case "pngblip":
                        Image.format = 'png';
                        break;
                    case "jpegblip":
                        Image.format = 'jpeg';
                        break;
                    case "macpict":
                        Image.format = 'pict';
                        break;
                    // case "wmetafile":
                    //     Image.format = 'bmp';
                    //     break;
    
                    // Picture size and scaling
                    case "picw":
                        Image.width = child.parameter;
                        break;
                    case "pich":
                        Image.height = child.parameter;
                        break;
                    case "picwgoal":
                        Image.goalWidth = child.parameter;
                        break;
                    case "pichgoal":
                        Image.goalHeight = child.parameter;
                        break;
                    case "picscalex":
                        Image.pcScaleX = child.parameter;
                        break;
                    case "picscaley":
                        Image.pcScaleY = child.parameter;
                        break;
    
                    // Binary or Hexadecimal Data ?
                    case "bin":
                        Image.binarySize = child.parameter;
                        break;
                    default:
                        break;
                }
            } else if (child instanceof RtfText) { // store Data
                Image.ImageData = child.text;
            }
        }
        // output Image
        this.output += Image.PrintImage();
        Image.Reset();
    }
    
    ProcessGroup(group) {
        // Can we ignore this group?
        switch (group.GetType()) {
            case "fonttbl": // Extract Font table
                this.ExtractFontTable(group.children);
                return;
            case "colortbl": // Extract color table
                this.ExtractColorTable(group.children);
                return;
            case "stylesheet":
                // Stylesheet extraction not yet supported
                return;
            case "info":
                // Ignore Document information
                return;
            case "pict":
                this.ExtractImage(group.children);
                return;
            case "nonshppict":
                // Ignore alternative images
                return;
            case "*": // Process destination
                this.ProcessDestination(group.children);
                return;
        }
    
        // Pictures extraction not yet supported
        // if (group.GetType().startsWith("pict")) return;
    
        // Push a new state onto the stack:
        // this.state = ;
        this.states.push(this.state);
    
        for (const child of group.children) {
            this.FormatEntry(child);
        }
    
        // Pop state from stack
        this.states.pop();
        this.state = this.states[this.states.length - 1];
    }
    
    ProcessDestination(dest) {
        if (!(dest[1] instanceof RtfControlWord)) return;
        // Check if this is a Word 97 picture
        if (dest[1].word === "shppict") {
            const c = dest.length;
            for (let i = 2; i < c; i++) {
                this.FormatEntry(dest[i]);
            }
        }
    }
    
    FormatEntry(entry) {
        if (entry instanceof RtfGroup) this.ProcessGroup(entry);
        else if (entry instanceof RtfControlWord) this.FormatControlWord(entry);
        else if (entry instanceof RtfControlSymbol) this.FormatControlSymbol(entry);
        else if (entry instanceof RtfText) this.FormatText(entry);
    }
    
    FormatControlWord(word) {
        // plain: Reset font formatting properties to default.
        // pard: Reset to default paragraph properties.
        if (word.word === "plain" || word.word === "pard") {
            this.state.Reset(this.defaultFont);
        } else if (word.word === "b") {
            this.state.bold = word.parameter; // bold
        } else if (word.word === "i") {
            this.state.italic = word.parameter; // italic
        } else if (word.word === "ul") {
            this.state.underline = word.parameter; // underline
        } else if (word.word === "ulnone") {
            this.state.underline = false; // no underline
        } else if (word.word === "strike") {
            this.state.strike = word.parameter; // strike through
        } else if (word.word === "v") {
            this.state.hidden = word.parameter; // hidden
        } else if (word.word === "fs") {
            this.state.fontsize = Math.ceil((word.parameter / 24) * 16); // font size
        } else if (word.word === "f") {
            this.state.font = word.parameter;
        }
        // Other formatting properties can be added here...
    
        // Unicode characters:
        else if (word.word === "u") {
            const uchar = this.DecodeUnicode(word.parameter);
            this.Write(uchar);
        }
        // More formatting properties can be added here...
    }
    
    // Other methods can be translated similarly...
    decodeUnicode(code, srcEnc = 'UTF-8') {
        let utf8 = '';
    
        if (srcEnc !== 'UTF-8') {
            utf8 = String.fromCharCode(iconv(srcEnc, 'UTF-8', code));
        }
    
        if (this.encoding === 'HTML-ENTITIES') {
            return utf8 ? `&#${ordUtf8(utf8)};` : `&#${code};`;
        } else if (this.encoding === 'UTF-8') {
            return utf8 ? utf8 : mbConvertEncoding(`&#${code};`, this.encoding, 'HTML-ENTITIES');
        } else {
            return utf8 ? mbConvertEncoding(utf8, this.encoding, 'UTF-8') :
                mbConvertEncoding(`&#${code};`, this.encoding, 'HTML-ENTITIES');
        }
    }
    
    write(txt) {
        if (!this.state.isLike(this.previousState) ||
            (this.state.isLike(this.previousState) && !this.openedTags['span'])) {
            this.CloseTag('span');
    
            const style = this.state.PrintStyle();
            this.previousState = { ...this.state };
    
            const attr = style ? `style="${style}"` : '';
            this.OpenTag('span', attr);
        }
        this.output += txt;
    }
    
    OpenTag(tag, attr = '') {
        this.output += attr ? `<${tag} ${attr}>` : `<${tag}>`;
        this.openedTags[tag] = true;
    }

    CloseTag(tag) {
        if (this.openedTags[tag]) {
            if (this.output.endsWith(`<${tag}>`)) {
                switch (tag) {
                    case 'p':
                        this.output = this.output.slice(0, -3) + "<br>";
                        break;
                    default:
                        this.output = this.output.slice(0, -(`</${tag}>`.length));
                        break;
                }
            } else {
                this.output += `</${tag}>`;
                this.openedTags[tag] = false;
            }
        }
    }
    
    CloseTags() {
        for (const tag in this.openedTags) {
            if (this.openedTags.hasOwnProperty(tag)) {
                this.CloseTag(tag);
            }
        }
    }
    
    FormatControlSymbol(symbol) {
        if (symbol.symbol === '\'') {
            const enc = this.getSourceEncoding();
            const uchar = this.decodeUnicode(symbol.parameter, enc);
            this.write(uchar);
        } else if (symbol.symbol === '~') {
            this.write("&nbsp;");
        } else if (symbol.symbol === '-') {
            this.write("&#173;");
        } else if (symbol.symbol === '_') {
            this.write("&#8209;");
        }
    }
    
    FormatText(text) {
        const txt = _.escape(text.text);
        if (this.encoding === 'HTML-ENTITIES') {
            this.write(txt);
        } else {
            this.write(mbConvertEncoding(txt, this.encoding, 'UTF-8'));
        }
    }

    GetSourceEncoding() {
        if (this.state.font) {
            if (RtfState.fonttbl[this.state.font]?.codepage) {
                return RtfState.fonttbl[this.state.font].codepage;
            } else if (RtfState.fonttbl[this.state.font]?.charset) {
                return RtfState.fonttbl[this.state.font].charset;
            }
        }
        return this.RTFencoding;
    }

    GetEncodingFromCharset(fcharset) {
        // Maps Windows character sets to encoding names
        const charset = {
            0: 'CP1252', // ANSI: Western Europe
            1: 'CP1252', // *Default
            2: 'CP1252', // *Symbol
            3: null,     // Invalid
            77: 'MAC',   // *also [MacRoman]: Macintosh
            128: 'CP932',  // *or [Shift_JIS]?: Japanese
            129: 'CP949',  // *also [UHC]: Korean (Hangul)
            130: 'CP1361', // *also [JOHAB]: Korean (Johab)
            134: 'CP936',  // *or [GB2312]?: Simplified Chinese
            136: 'CP950',  // *or [BIG5]?: Traditional Chinese
            161: 'CP1253', // Greek
            162: 'CP1254', // Turkish (latin 5)
            163: 'CP1258', // Vietnamese
            177: 'CP1255', // Hebrew
            178: 'CP1256', // Simplified Arabic
            179: 'CP1256', // *Traditional Arabic
            180: 'CP1256', // *Arabic User
            181: 'CP1255', // *Hebrew User
            186: 'CP1257', // Baltic
            204: 'CP1251', // Russian (Cyrillic)
            222: 'CP874',  // Thai
            238: 'CP1250', // Eastern European (latin 2)
            254: 'CP437',  // *also [IBM437][437]: PC437
            255: 'CP437',  // *OEM still PC437
        };
    
        if (charset[fcharset] !== undefined) {
            return charset[fcharset];
        } else {
            console.error(`Unknown charset: ${fcharset}`);
        }
    }

    etEncodingFromCodepage(cpg) {
        const codePage = {
            'ansi': 'CP1252',
            'mac': 'MAC',
            'pc': 'CP437',
            'pca': 'CP850',
            437: 'CP437', // United States IBM
            708: 'ASMO-708', // also [ISO-8859-6][ARABIC] Arabic
            // Not supported by iconv
            // 709: '', // Arabic (ASMO 449+, BCON V4)
            // 710: '', // Arabic (transparent Arabic)
            // 711: '', // Arabic (Nafitha Enhanced)
            // 720: '', // Arabic (transparent ASMO)
            819: 'CP819', // Windows 3.1 (US and Western Europe)
            850: 'CP850', // IBM multilingual
            852: 'CP852', // Eastern European
            860: 'CP860', // Portuguese
            862: 'CP862', // Hebrew
            863: 'CP863', // French Canadian
            864: 'CP864', // Arabic
            865: 'CP865', // Norwegian
            866: 'CP866', // Soviet Union
            874: 'CP874', // Thai
            932: 'CP932', // Japanese
            936: 'CP936', // Simplified Chinese
            949: 'CP949', // Korean
            950: 'CP950', // Traditional Chinese
            1250: 'CP1250', // Windows 3.1 (Eastern European)
            1251: 'CP1251', // Windows 3.1 (Cyrillic)
            1252: 'CP1252', // Western European
            1253: 'CP1253', // Greek
            1254: 'CP1254', // Turkish
            1255: 'CP1255', // Hebrew
            1256: 'CP1256', // Arabic
            1257: 'CP1257', // Baltic
            1258: 'CP1258', // Vietnamese
            1361: 'CP1361' // Johab
        };
    
        if (codePage[cpg] !== undefined) {
            return codePage[cpg];
        } else {
            console.error(`Unknown codepage: ${cpg}`);
        }
    }

    ord_utf8(chr) {
        const ord0 = chr.charCodeAt(0);
        if (ord0 >= 0 && ord0 <= 127) {
            return ord0;
        }
        const ord1 = chr.charCodeAt(1);
        if (ord0 >= 192 && ord0 <= 223) {
            return (ord0 - 192) * 64 + (ord1 - 128);
        }
        const ord2 = chr.charCodeAt(2);
        if (ord0 >= 224 && ord0 <= 239) {
            return (ord0 - 224) * 4096 + (ord1 - 128) * 64 + (ord2 - 128);
        }
        const ord3 = chr.charCodeAt(3);
        if (ord0 >= 240 && ord0 <= 247) {
            return (ord0 - 240) * 262144 + (ord1 - 128) * 4096 + (ord2 - 128) * 64 + (ord3 - 128);
        }
        const ord4 = chr.charCodeAt(4);
        if (ord0 >= 248 && ord0 <= 251) {
            return (ord0 - 248) * 16777216 + (ord1 - 128) * 262144 + (ord2 - 128) * 4096 + (ord3 - 128) * 64 + (ord4 - 128);
        }
        if (ord0 >= 252 && ord0 <= 253) {
            return (ord0 - 252) * 1073741824 + (ord1 - 128) * 16777216 + (ord2 - 128) * 262144 + (ord3 - 128) * 4096 + (ord4 - 128) * 64 + (chr.charCodeAt(5) - 128);
        }
        console.error(`Invalid Unicode character: ${chr}`);
    }        
}

const fs = require('fs');

// Read the RTF file content
const rtfFilePath = 'test.rtf';
const rtfContent = fs.readFileSync(rtfFilePath, 'utf-8');

// Parse the RTF content
const reader = new RtfReader();
reader.Parse(rtfContent);

// Format RTF content to HTML
const formatter = new RtfHtml();
const formattedString = formatter.Format(reader.root).trim();

// Write formatted HTML content to a file
const htmlFilePath = 'some.html';
fs.writeFileSync(htmlFilePath, formattedString, 'utf-8');

console.log('Conversion from RTF to HTML complete.');