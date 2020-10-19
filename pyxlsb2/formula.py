from .tokenreader import TokenReader


class Formula(object):
    def __init__(self, tokens):
        self._tokens = list(tokens)

    def __repr__(self):
        return 'Formula({})'.format(self._tokens)

    def __str__(self):
        return self.stringify()

    def stringify(self, workbook):
        tokens = self._tokens[:]
        return '' if not tokens else tokens.pop().stringify(tokens, workbook)

    @classmethod
    def parse(cls, data):
        return cls(TokenReader(data))
