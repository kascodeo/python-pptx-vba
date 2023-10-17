from collections import OrderedDict, namedtuple
import re


class TextRange():
    def __init__(self, textframe, Start=None, Length=None):
        self._textframe = textframe
        self._start = Start
        self._length = Length

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
        Start = self.get_start()
        Length = self.get_length()

        text = ''
        prev_p = None
        chars = self.get_info()['chars']
        for i in range(Start-1, Start-1 + Length):
            ch = chars[i]
            p = ch.p
            if prev_p is not None and prev_p is not p:
                text += '\r\n'
            prev_p = p
            text += ch.c
        return text

    @Text.setter
    def Text(self, NewText):
        Start = self.get_start()
        Length = self.get_length()

        flag = False
        if Start == 1 and self._length is None:
            self.add_run_if_not_present()
            p = self.e.findqn('a:p')
            t = p.findqn('a:r').findqn('a:t')
            t.text = ''
            flag = True

        self.split_p_after(Start+Length-1)

        # for splitting p after Start -1
        flag = False
        if Start == 1:
            t = self._t(self.get_info()['chars'][0].r)
            t.text = 'Z'+t.text
            Start += 1
            flag = True

        self.split_p_after(Start-1)

        if flag:
            t = self._t(self.get_info()['chars'][0].r)
            t.text = t.text[1:]
            Start -= 1
        # end of splitting

        chars = self.get_info()['chars']
        p_lst = []
        for i in range(Start-1, Start-1 + Length):
            ch = chars[i]
            if ch.p not in p_lst:
                p_lst.append(ch.p)

        # remove paragraphs from second to end in range
        for p in p_lst[1:]:
            p.getparent().remove(p)

        # remove runs except first, in first para
        p = p_lst[0]
        r_lst = p.findallqn('a:r')
        for r in r_lst[1:]:
            r.getparent().remove(r)
        r = r_lst[0]
        self._t(r).text = ''
        if self._length is not None:
            self._length = self.Characters(Start, 1).InsertBefore(
                NewText)._length

    def Characters(self, Start, Length):
        return self.__class__(self._textframe, Start, Length)

    def get_info(self):
        chars = OrderedDict()
        Char = namedtuple('Char', ['ip', 'ir',  'ic', 'p', 'r', 'c', 'irp'])
        ic = -1
        for p in self.e.findallqn('a:p'):
            ip = -1
            for irp, r in enumerate(p.findallqn('a:r')):
                t = r.findqn('a:t')
                if t is not None:
                    for ir, c in enumerate(t.text):
                        ic += 1
                        ip += 1
                        chars[ic] = Char(ip, ir, ic, p, r, c, irp)
        return {'chars': chars}

    def InsertAfter(self, NewText):
        if len(NewText) == 0 or not isinstance(NewText, str):
            raise ValueError("NewText must be non zero length string")
        Start = self.get_start()
        Length = self.get_length()

        self._insert_after(NewText, Start, Length)

        newlength = self._length_of_text(NewText)
        return self.__class__(self._textframe, Start+Length, newlength)

    def _length_of_text(self, NewText):
        return sum([len(line) for line in self.get_lines(NewText)])

    def _insert_after(self, NewText, Start, Length):
        lines = self.get_lines(NewText)

        End = Start + Length - 1

        # splitting the para needed because additional paras to be inserted
        if len(lines) > 1:
            self.split_p_after(End)

        # insert the first line in given text at end of the end char
        line = lines[0]
        chars = self.get_info()['chars']
        char = chars[End-1]
        t = self._t(char.r)
        t.text = t.text[:char.ir+1]+line+t.text[char.ir+1:]

        # create new para each for second line onwards after the above para
        p = char.p
        for line in lines[1:]:
            p_ = p.deepcopy()
            p.addnext(p_)
            for r_ in p_.findallqn('a:r')[1:]:  # remove all r but first
                r_.getparent().remove(r_)
            r_ = p_.findqn('a:r')
            t_ = self._t(r_)
            t_.text = line  # replace the text with the given line text
            p = p_  # needed to add the next para after the current added para

    def InsertBefore(self, NewText):
        if len(NewText) == 0 or not isinstance(NewText, str):
            raise ValueError("NewText must be non zero length string")

        Start = self.get_start()

        flag = False
        if Start == 1:
            t = self._t(self.get_info()['chars'][0].r)
            t.text = 'Z'+t.text
            Start += 1
            flag = True

        tr_previous_char = self.Characters(Start-1, 1)
        tr_previous_char.InsertAfter(NewText)

        if flag:
            t = self._t(self.get_info()['chars'][0].r)
            t.text = t.text[1:]
            Start -= 1

        newlength = self._length_of_text(NewText)
        return self.__class__(self._textframe, Start, newlength)

    def split_p_after(self, End):
        if self.is_char_at_end_of_p(End):
            return  # end char is already at end of para. no need to split

        # split the run after the end char
        self.split_r_after(End)

        # gather all runs after the end char in the para
        chars = self.get_info()['chars']
        char = chars[End-1]
        p = char.p
        r_lst = []
        for i in range(char.ic+1, len(chars)):
            if p is chars[i].p:
                if chars[i].r not in r_lst:
                    r_lst.append(chars[i].r)

        # copy the para add next, remove all runs, move fathered runs to para
        p_ = p.deepcopy()
        p.addnext(p_)
        for r in p_.findallqn('a:r'):
            r.getparent().remove(r)
        for i, r in enumerate(r_lst):
            p_.insert(i, r)

        #

    def split_r_after(self, End):
        if self.is_char_at_end_of_r(End):
            return  # end char is already at end of run. no need to split

        chars = self.get_info()['chars']
        char = chars[End-1]
        r = char.r
        t = self._t(r)
        r_ = r.deepcopy()   # copy and add next to the run
        t.text = t.text[:char.ir+1]  # keep chars till end char
        r.addnext(r_)
        t_ = self._t(r_)
        t_.text = t_.text[char.ir+1:]   # keep chars after end char

    def is_char_at_end_of_r(self, End):
        chars = self.get_info()['chars']
        char = chars[End-1]
        if (char.ic+1 not in chars) or not (char.r is chars[char.ic+1].r):
            return True
        return False

    def is_char_at_end_of_p(self, End):
        chars = self.get_info()['chars']
        char = chars[End-1]
        if (char.ic + 1 not in chars) or not (char.p is chars[char.ic + 1].p):
            return True
        return False

    def add_run_if_not_present(self, p=None):
        if p is None:
            p = self.e.findqn('a:p')
        if p is not None and p.findqn('a:r') is None:
            pr = p.findqn('a:endParaRPr')
            attrib = {}
            if pr is not None:
                attrib = pr.attrib
            r = p.makeelement(p.qn('a:r'), attrib=attrib)
            p.insert(0, r)
            t = r.makeelement(r.qn('a:t'))
            r.insert(0, t)

    def get_lines(self, txt):
        return re.split(r'\r\n|\r|\n', txt)

    def get_start(self):
        return 1 if self._start is None else self._start

    def get_length(self):
        return len(self.get_info()['chars']) if self._length is None \
            else self._length

    def _t(self, r):
        return r.findqn('a:t')

    def get_chars(self):
        Start = self.get_start()
        Length = self.get_length()

        chars = self.get_info()['chars']
        lst = []
        for i in range(Start-1, Start-1 + Length):
            lst.append(chars[i])
        return lst

    def get_runs(self):
        lst = []
        for ch in self.get_chars():
            if ch.r not in lst:
                lst.append(ch.r)
        return lst

    def get_rpr(self):
        lst = []
        for r in self.get_runs():
            rpr = r.findqn('a:rPr')
            lst.append(rpr)
        return lst

    def isolate(self):
        Start = self.get_start()
        Length = self.get_length()
        End = Start + Length - 1

        self.split_p_after(End)
        if Start > 1:
            self.split_p_after(Start-1)
