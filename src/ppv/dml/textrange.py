from collections import OrderedDict, namedtuple
import re


class TextRange():
    def __init__(self, textframe, istart=None, length=None):
        self._textframe = textframe
        self._istart = istart
        self._length = length

    @property
    def Parent(self):
        return self._textframe

    @property
    def e(self):
        return self.Parent.e

    @property
    def Font(self):
        from ppv.dml.font import Font
        return Font(self)

    @property
    def Text(self):
        text = ''
        prev_p = None
        chars = self.get_chars_in_range()
        for i, ch in enumerate(chars):
            p = ch.p
            if prev_p is not None and prev_p is not p:
                text += '\r\n'
            prev_p = p
            text += ch.c
        return text

    @Text.setter
    def Text(self, NewText):
        # isolate
        self.isolate()
        chars = self.get_chars_in_range()

        lst_r = []
        for ch in chars:
            r = ch.r
            if r not in lst_r:
                lst_r.append(r)
        lst_p = []
        for ch in chars:
            p = ch.p
            if p not in lst_p:
                lst_p.append(p)

        for r in lst_r:
            r.getparent().remove(r)

        for p in lst_p[1:-1]:
            p.getparent().remove(p)

        if len(lst_p) > 1:
            self.join_para(lst_p[0], lst_p[-1])

        if NewText == '':
            return

        if len(chars) == 0:
            if self.istart != 0:
                raise ValueError("Use new TextRange.")
            else:
                tr = self.Parent.TextRange
                tr.InsertBefore(NewText)
        else:
            tr = self.__class__(self._textframe, self.istart-1, 1)
            tr.InsertAfter(NewText)

    def InsertAfter(self, NewText):
        if len(NewText) == 0 or not isinstance(NewText, str):
            raise ValueError("NewText must be non zero length string")

        self.split_right()
        lines = self.get_lines(NewText)
        line_first = lines[0]
        chars = self.get_chars_in_range()

        flag = False
        if len(chars) == 0:
            if self.istart == 0:
                p, r, t = self.get_or_create_p_r_t()
                t.text = 'Z'
                flag = True
                chars = self.get_chars_in_range()
            else:
                raise ValueError("zero Length. Use other Text Range object")

        ch = chars[-1]
        r = ch.r
        p = ch.p
        if line_first != '':
            r_ = r.deepcopy()
            r.addnext(r_)
            t_ = r_.findqn('a:t')
            t_.text = line_first
            r = r_
        if len(lines) > 1:
            lst_r = []
            for e in r.itersiblings():
                if e.tag == r.tag:
                    lst_r.append(e)
            p_ = self.add_para(p)
            for r in reversed(lst_r):
                p_.insert(0, r)

        p = ch.p
        for line in lines[1:]:
            p = self.add_para(p)
            r, rpr, t = self.add_run(p)
            t.text = line
        p_ = p.getnext()
        if p_ is not None and p_.tag == p.tag:
            self.join_para(p, p_)

        if flag:
            t = chars[0].t
            text = self.get_text(t)
            t.text = text[1:]
            self._del_r_if_empty(t)

    def InsertBefore(self, NewText):
        if len(NewText) == 0 or not isinstance(NewText, str):
            raise ValueError("NewText must be non zero length string")

        if self.istart - 1 in self.get_chars():
            tr = self.__class__(self._textframe, self.istart-1, 1)
            tr.InsertAfter(NewText)
        else:
            p, r, t = self.get_or_create_p_r_t()
            text = self.get_text(t)
            t.text = 'Z'+text
            tr = self.__class__(self._textframe, 0, 1)
            tr.InsertAfter(NewText)
            t.text = t.text[1:]
            self._del_r_if_empty(t)

# -----------------------------------------------

    def _del_r_if_empty(self, t):
        if t.text == '' or t.text is None:
            r = t.getparent()
            r.getparent().remove(r)

    def get_chars(self):
        chars = OrderedDict()
        Char = namedtuple('Char',
                          ['ip', 'ir', 'ic', 'p', 'r', 't', 'c', 'irp'])
        ic = -1
        for p in self.e.findallqn('a:p'):
            ip = -1
            for irp, r in enumerate(p.findallqn('a:r')):
                t = r.findqn('a:t')
                if t is not None and t.text is not None:
                    for ir, c in enumerate(t.text):
                        ic += 1
                        ip += 1
                        chars[ic] = Char(ip, ir, ic, p, r, t, c, irp)
        return chars

    def get_chars_in_range(self):
        chars = self.get_chars()
        lst = []
        for i in range(self.istart, self.iend+1):
            char = chars[i]
            lst.append(char)
        return lst

    @property
    def istart(self):
        if self._istart is None:
            return 0
        if self._istart < 0:
            return 0
        if self._istart > len(self.get_chars()) - 1:
            return len(self.get_chars()) - 1
        return self._istart

    @property
    def length(self):
        if self._length is None:
            return len(self.get_chars())
        return self._length

    @property
    def iend(self):
        res = self.istart + self.length - 1
        # if res < 0:
        #     return 0
        if res > len(self.get_chars())-1:
            return len(self.get_chars()) - 1
        return res

    def isolate(self):
        self.split_left()
        self.split_right()

    def split_left(self):
        chars = self.get_chars_in_range()
        if len(chars) == 0:
            return

        char = chars[0]
        if char.ir == 0:
            return

        r = char.r
        r_ = r.deepcopy()
        r.addnext(r_)
        t = char.t
        text = self.get_text(t)
        t.text = text[:char.ir]
        t_ = r_.findqn('a:t')
        text_ = self.get_text(t_)
        t_.text = text_[char.ir:]

    def split_right(self):
        chars = self.get_chars_in_range()
        if len(chars) == 0:
            return

        char = chars[-1]
        if char.ir == len(char.t.text)-1:
            return

        r = char.r
        r_ = r.deepcopy()
        r.addnext(r_)
        t = char.t
        t.text = t.text[:char.ir+1]
        t_ = r_.findqn('a:t')
        t_.text = t_.text[char.ir+1:]

    def join_para(self, base_p, del_p):
        endpara = del_p[-1]
        endpara.getparent().remove(endpara)
        for e in del_p:
            base_p.append(e)
        del_p.getparent().remove(del_p)

    def get_or_create_para(self):
        p = self.e.findqn('a:p')
        if p is None:
            p = self.e.makeelement(self.e.qn('a:p'))
            self.e.append(p)
            p.append(p.makeelement(p.qn('a:endParaRPr')))
        return p

    def add_para(self, after_p):
        p = after_p.deepcopy()
        after_p.addnext(p)
        for r in p.findallqn('a:r'):
            r.getparent().remove(r)
        return p

    def add_run(self, p):
        r = p.makeelement(p.qn('a:r'), attrib=p[-1].attrib)
        t = p.makeelement(p.qn('a:t'))
        p.insert(-1, r)

        rpr = r.makeelement(r.qn('a:rPr'))
        r.insert(0, rpr)
        r.insert(1, t)
        return r, rpr, t

    def get_lines(self, txt):
        return re.split(r'\r\n|\r|\n', txt)

    def get_or_create_p_r_t(self):
        p = self.get_or_create_para()
        r = p.findqn('a:r')
        if r is None:
            r, rpr, t = self.add_run(p)
        t = r.findqn('a:t')
        return p, r, t

    def get_text(self, t):
        if t.text is None:
            return ''
        return t.text
