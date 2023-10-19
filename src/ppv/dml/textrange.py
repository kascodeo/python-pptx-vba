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

    def Characters(self, Start=-1, Length=-1):
        if Start == 0:
            Start = 1
        if Length == 0:
            Length = 1

        lst_c = self.get_c_in_range()
        num = len(lst_c)
        if Start == -1 and Length == -1:
            # Both Start and Length are Omitted then returned range is of all
            # paragraphs
            start_index = 0
            end_index = start_index + num - 1

        if Start != -1 and Length == -1:
            # Just Start is specified then returned range is specified
            # paragraph only
            start_index = Start - 1
            end_index = Start - 1

        if Start == -1 and Length != -1:
            # Start is Omitted and Length is specified then
            # returned range starts from first paragraph till length of
            # paragraphs
            start_index = 0
            end_index = start_index + Length - 1

        if Start != -1 and Length != -1:
            # Both Start and Length are specified
            start_index = Start - 1
            end_index = start_index + Length - 1

        if start_index > num - 1:
            start_index = num - 1

        if end_index > num - 1:
            end_index = num - 1

        lst = lst_c[start_index:end_index+1]
        c_start = lst[0]
        c_end = lst[-1]
        istart = c_start.ic
        iend = c_end.ic

        return self.__class__(self._textframe, istart, iend-istart+1)

    def Paragraphs(self, Start=-1, Length=-1):
        if Start == 0:
            Start = 1
        if Length == 0:
            Length = 1

        lst_p = self.get_p_in_range()
        num = len(lst_p)
        if Start == -1 and Length == -1:
            # Both Start and Length are Omitted then returned range is of all
            # paragraphs
            start_index = 0
            end_index = start_index + num - 1

        if Start != -1 and Length == -1:
            # Just Start is specified then returned range is specified
            # paragraph only
            start_index = Start - 1
            end_index = Start - 1

        if Start == -1 and Length != -1:
            # Start is Omitted and Length is specified then
            # returned range starts from first paragraph till length of
            # paragraphs
            start_index = 0
            end_index = start_index + Length - 1

        if Start != -1 and Length != -1:
            # Both Start and Length are specified
            start_index = Start - 1
            end_index = start_index + Length - 1

        if start_index > num - 1:
            start_index = num - 1

        if end_index > num - 1:
            end_index = num - 1

        lst = lst_p[start_index:end_index+1]
        p_start = lst[0]
        p_end = lst[-1]
        istart = 0
        iend = 0
        flag_istart = True
        chars = self.get_chars()
        for char in chars.values():
            p = char.p
            if flag_istart and p is p_start:
                istart = char.ic
                flag_istart = False
            if p is p_end:
                iend = char.ic

        return self.__class__(self._textframe, istart, iend-istart+1)

    def Runs(self, Start=-1, Length=-1):
        if Start == 0:
            Start = 1
        if Length == 0:
            Length = 1

        lst_r = self.get_r_in_range()
        num = len(lst_r)
        if Start == -1 and Length == -1:
            # Both Start and Length are Omitted then returned range is of all
            # paragraphs
            start_index = 0
            end_index = start_index + num - 1

        if Start != -1 and Length == -1:
            # Just Start is specified then returned range is specified
            # paragraph only
            start_index = Start - 1
            end_index = Start - 1

        if Start == -1 and Length != -1:
            # Start is Omitted and Length is specified then
            # returned range starts from first paragraph till length of
            # paragraphs
            start_index = 0
            end_index = start_index + Length - 1

        if Start != -1 and Length != -1:
            # Both Start and Length are specified
            start_index = Start - 1
            end_index = start_index + Length - 1

        if start_index > num - 1:
            start_index = num - 1

        if end_index > num - 1:
            end_index = num - 1

        lst = lst_r[start_index:end_index+1]
        r_start = lst[0]
        r_end = lst[-1]
        istart = 0
        iend = 0
        flag_istart = True
        chars = self.get_chars()
        for char in chars.values():
            r = char.r
            if flag_istart and r is r_start:
                istart = char.ic
                flag_istart = False
            if r is r_end:
                iend = char.ic

        return self.__class__(self._textframe, istart, iend-istart+1)

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

        if self._length is not None:
            self._length = self.get_length_lines(self.get_lines(NewText))

    def InsertAfter(self, NewText):
        if len(NewText) == 0 or not isinstance(NewText, str):
            raise ValueError("NewText must be non zero length string")

        iend = self.iend
        self.split_right()
        lines = self.get_lines(NewText)
        line_first = lines[0]
        chars = self.get_chars_in_range()

        flag = False
        if len(chars) == 0:

            if self.istart == 0:
                # add a char in this range will be removed later
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
            # add a new run after the run of last char of this range
            r_ = r.deepcopy()
            r.addnext(r_)
            t_ = r_.findqn('a:t')
            t_.text = line_first  # first line is the text of new run
            r = r_  # new run is the reference for the coming runs

        if len(lines) > 1:
            # move the remaining runs in the above para to new para
            # this is to move text after current range to new para
            # which is not needed when NewText has no para in it
            lst_r = []
            for e in r.itersiblings():
                if e.tag == r.tag:
                    lst_r.append(e)
            p_ = self.add_para(p)

            endParaRPr = r.getnext()
            if endParaRPr.ln == 'endParaRPr':
                endParaRPr.getparent().remove(endParaRPr)
                rpr_ = r.findqn('a:rPr')
                endParaRPr = rpr_.deepcopy()
                r.addnext(endParaRPr)

            for r in reversed(lst_r):
                p_.insert(0, r)

        p = ch.p
        for line in lines[1:]:
            # NewText has new paras
            # for each new para add para from first para
            # add new run in new para and add the text
            p = self.add_para(p)
            r, rpr, t = self.add_run(p)
            t.text = line
        p_ = p.getnext()
        if p_ is not None and p_.tag == p.tag:
            # join
            self.join_para(p, p_)

        if flag:
            # remove the dummy text if added before
            t = chars[0].t
            text = self.get_text(t)
            t.text = text[1:]
            self._del_r_if_empty(t)

        tr_new = self.__class__(self._textframe, iend+1,
                                self.get_length_lines(lines))
        return tr_new

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

        lines = self.get_lines(NewText)
        tr_new = self.__class__(self._textframe, self.istart,
                                self.get_length_lines(lines))
        if self._istart is not None:
            self._istart = self._istart + self.get_length_lines(lines)
        return tr_new
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
        endpara = base_p[-1]
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
        # add para with no runs but same endParaRPr.
        # para is added next to after_p
        p = after_p.deepcopy()
        after_p.addnext(p)
        for r in p.findallqn('a:r'):
            r.getparent().remove(r)
        return p

    def add_run(self, p):
        r = p.makeelement(p.qn('a:r'))
        t = p.makeelement(p.qn('a:t'))
        p.insert(-1, r)

        endParaRPr = p[-1]
        rpr = endParaRPr.deepcopy()
        rpr.tag = rpr.qn('a:rPr')
        r.insert(0, rpr)
        r.insert(1, t)
        return r, rpr, t

    def get_lines(self, txt):
        return re.split(r'\r\n|\r|\n', txt)

    def get_length_lines(self, lines):
        count = 0
        for line in lines:
            count += len(line)
        return count

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

    def get_p_in_range(self):
        chars = self.get_chars_in_range()
        lst = []
        for char in chars:
            p = char.p
            if p not in lst:
                lst.append(p)
        return lst

    def get_r_in_range(self):
        chars = self.get_chars_in_range()
        lst = []
        for char in chars:
            r = char.r
            if r not in lst:
                lst.append(r)
        return lst

    def get_c_in_range(self):
        return self.get_chars_in_range()

    def get_rpr_in_range(self):
        return [r.findqn('a:rPr') for r in self.get_r_in_range()]
        # chars = self.get_chars_in_range()
        # lst = []
        # for char in chars:
        #     r = char.r
        #     rpr = r.findqn('a:rPr')
        #     if (rpr is not None) and (not (rpr in lst)):
        #         lst.append(rpr)
        # return lst

    def update_rpr(self, rpr, attr, value):
        a_latin = 'a:latin'
        endParaRPr = rpr.getparent().getnext()
        if endParaRPr is not None and endParaRPr.ln != 'endParaRPr':
            endParaRPr = None

        if value is None:
            if attr == 'font':
                # for font name delete
                latin = rpr.findqn(a_latin)
                if latin is not None:
                    latin.getparent().remove(latin)

                # delete font name for endParaRPr also if run is last in para
                if endParaRPr is not None:
                    latin = endParaRPr.findqn(a_latin)
                    if latin is not None:
                        latin.getparent().remove(latin)
            else:
                # delete bold, underline etc
                if attr in rpr.attrib:
                    del rpr.attrib[attr]
                # delete from endpararpr also if run is last in para
                if endParaRPr is not None:
                    if attr in endParaRPr.attrib:
                        del endParaRPr.attrib[attr]
        else:
            if attr == 'font':
                # for font name update
                latin = rpr.findqn(a_latin)
                if latin is None:
                    latin = rpr.makeelement(rpr.qn(a_latin))
                    rpr.append(latin)
                latin.attrib['typeface'] = value

                # update in the endpararpr also if run is last in para
                if endParaRPr is not None:
                    latin = endParaRPr.findqn(a_latin)
                    # create latin if not present
                    if latin is None:
                        latin = endParaRPr.makeelement(rpr.qn(a_latin))
                        endParaRPr.append(latin)
                    latin.attrib['typeface'] = value

            else:
                # for bold underline etc update
                rpr.attrib[attr] = value
                if endParaRPr is not None:
                    endParaRPr.attrib[attr] = value
