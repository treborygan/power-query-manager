"""Small M/diff editor with line numbers and editor-local search."""
from PySide6.QtCore import QEvent, QRect, QSize, Qt
from PySide6.QtGui import (QColor, QFont, QPainter, QPalette, QTextCharFormat, QTextCursor,
                           QTextDocument, QSyntaxHighlighter)
from PySide6.QtWidgets import QPlainTextEdit, QTextEdit, QWidget


class MCodeHighlighter(QSyntaxHighlighter):
    """Readable token coloring for common M syntax; this is not an M parser."""
    def __init__(self, document):
        super().__init__(document)
        self.set_theme(False)

    def set_theme(self,dark):
        if dark:
            colors={'keyword':'#c4a1ff','string':'#67dfa0','comment':'#aebccc',
                    'number':'#ffb86b','identifier':'#77c4ff'}
        else:
            colors={'keyword':'#7c3aed','string':'#087443','comment':'#596579',
                    'number':'#b45309','identifier':'#125f9c'}
        self.keyword_format=self._format(colors['keyword'],bold=True)
        self.string_format=self._format(colors['string'])
        self.comment_format=self._format(colors['comment'],italic=True)
        self.number_format=self._format(colors['number'])
        self.identifier_format=self._format(colors['identifier'],bold=True)
        self.rehighlight()

    @staticmethod
    def _format(color,*,bold=False,italic=False):
        result=QTextCharFormat();result.setForeground(QColor(color))
        if bold:result.setFontWeight(QFont.Bold)
        if italic:result.setFontItalic(True)
        return result

    def highlightBlock(self,text):
        keywords={'let','in','if','then','else','each','and','or','not','as','is','error','try','otherwise','meta','type'}
        i=0;length=len(text)
        while i<length:
            if text.startswith('//',i):
                self.setFormat(i,length-i,self.comment_format);break
            if text[i]=='"':
                start=i;i+=1
                while i<length:
                    if text[i]=='"':
                        if i+1<length and text[i+1]=='"':i+=2;continue
                        i+=1;break
                    i+=1
                self.setFormat(start,i-start,self.string_format);continue
            if text[i].isdigit():
                start=i;i+=1
                while i<length and (text[i].isdigit() or text[i] in '.eE+-'):
                    if text[i] in '+-' and text[i-1] not in 'eE':break
                    i+=1
                self.setFormat(start,i-start,self.number_format);continue
            if text[i].isalpha() or text[i]=='_':
                start=i;i+=1
                while i<length and (text[i].isalnum() or text[i] in '_.'):i+=1
                word=text[start:i]
                if word.casefold() in keywords:self.setFormat(start,i-start,self.keyword_format)
                elif i<length and text[i]=='(':self.setFormat(start,i-start,self.identifier_format)
                continue
            i+=1


class DiffHighlighter(QSyntaxHighlighter):
    def __init__(self,document):
        super().__init__(document)
        self.set_theme(False)

    def set_theme(self,dark):
        if dark:
            colors={'added':('#76d49a','#193a2a'),'removed':('#ff9999','#422326'),
                    'header':('#a9cbff','#202f42')}
        else:
            colors={'added':('#176b3a','#e8f5ec'),'removed':('#a32929','#fceaea'),
                    'header':('#315c91','#edf3fb')}
        self.added=self._format(*colors['added'])
        self.removed=self._format(*colors['removed'])
        self.header=self._format(*colors['header'])
        self.rehighlight()

    @staticmethod
    def _format(foreground,background):
        result=QTextCharFormat();result.setForeground(QColor(foreground));result.setBackground(QColor(background))
        return result

    def highlightBlock(self,text):
        if text.startswith('+++ ') or text.startswith('--- '):self.setFormat(0,len(text),self.header)
        elif text.startswith('+'):self.setFormat(0,len(text),self.added)
        elif text.startswith('-'):self.setFormat(0,len(text),self.removed)
        elif text.startswith('@@'):self.setFormat(0,len(text),self.header)


class _LineNumberArea(QWidget):
    def __init__(self,editor):
        super().__init__(editor);self.editor=editor
    def sizeHint(self):return QSize(self.editor.lineNumberAreaWidth(),0)
    def _text_colors(self):
        dark=self.editor.palette().color(QPalette.Base).lightness()<128
        return ('#aab8c8','#242d37') if dark else ('#56657a','#f0f3f7')
    def paintEvent(self,event):self.editor.lineNumberAreaPaintEvent(event)


class CodeEditor(QPlainTextEdit):
    """Text editor with a numbered gutter, M/diff colors, and wraparound find."""
    def __init__(self,parent=None,*,mode='m'):
        super().__init__(parent)
        if mode not in {'m','diff'}:raise ValueError('Editor mode must be m or diff.')
        self.mode=mode
        self.highlighter=(MCodeHighlighter(self.document()) if mode=='m' else DiffHighlighter(self.document()))
        self.lineNumberArea=_LineNumberArea(self)
        self.blockCountChanged.connect(self._update_line_number_width)
        self.updateRequest.connect(self._update_line_number_area)
        self.cursorPositionChanged.connect(self._highlight_current_line)
        self._update_line_number_width(0);self._highlight_current_line()
        self.setLineWrapMode(QPlainTextEdit.NoWrap)
        self._update_theme_colors()

    def _update_theme_colors(self):
        dark=self.palette().color(QPalette.Base).lightness()<128
        self.highlighter.set_theme(dark)
        self.lineNumberArea.update()
        self._highlight_current_line()

    def changeEvent(self,event):
        super().changeEvent(event)
        if event.type() in (QEvent.Type.PaletteChange,QEvent.Type.StyleChange) and hasattr(self,'highlighter'):
            self._update_theme_colors()

    def lineNumberAreaWidth(self):
        digits=max(2,len(str(max(1,self.blockCount()))))
        return 8+self.fontMetrics().horizontalAdvance('9')*digits

    def _update_line_number_width(self,_=0):
        self.setViewportMargins(self.lineNumberAreaWidth(),0,0,0)

    def _update_line_number_area(self,rect,dy):
        if dy:self.lineNumberArea.scroll(0,dy)
        else:self.lineNumberArea.update(0,rect.y(),self.lineNumberArea.width(),rect.height())
        if rect.contains(self.viewport().rect()):self._update_line_number_width()

    def resizeEvent(self,event):
        super().resizeEvent(event)
        rect=self.contentsRect()
        self.lineNumberArea.setGeometry(QRect(rect.left(),rect.top(),self.lineNumberAreaWidth(),rect.height()))

    def lineNumberAreaPaintEvent(self,event):
        foreground,background=self.lineNumberArea._text_colors()
        painter=QPainter(self.lineNumberArea);painter.fillRect(event.rect(),QColor(background))
        block=self.firstVisibleBlock();number=block.blockNumber()
        top=int(self.blockBoundingGeometry(block).translated(self.contentOffset()).top())
        bottom=top+int(self.blockBoundingRect(block).height())
        painter.setPen(QColor(foreground))
        while block.isValid() and top<=event.rect().bottom():
            if block.isVisible() and bottom>=event.rect().top():
                painter.drawText(0,top,self.lineNumberArea.width()-5,self.fontMetrics().height(),
                                 Qt.AlignRight,str(number+1))
            block=block.next();top=bottom;bottom=top+int(self.blockBoundingRect(block).height());number+=1

    def _highlight_current_line(self):
        if self.isReadOnly():return
        dark=self.palette().color(QPalette.Base).lightness()<128
        selection=QTextEditExtraSelection();selection.format.setBackground(QColor('#3d3520' if dark else '#fffde7'))
        selection.format.setProperty(QTextCharFormat.FullWidthSelection,True)
        selection.cursor=self.textCursor();selection.cursor.clearSelection()
        self.setExtraSelections([selection])

    def find_text(self,text,*,backward=False):
        if not text:return False
        flags=QTextDocument.FindFlag.FindBackward if backward else QTextDocument.FindFlag(0)
        if self.find(text,flags):return True
        cursor=self.textCursor();cursor.movePosition(QTextCursor.End if backward else QTextCursor.Start)
        self.setTextCursor(cursor)
        return self.find(text,flags)


QTextEditExtraSelection=QTextEdit.ExtraSelection
