from .ptgs import NamePtg
from .tokenreader import TokenReader


class Formula(object):
    def __init__(self, tokens):
        self._tokens = list(tokens)

    def __repr__(self):
        return 'Formula({})'.format(self._tokens)

    def __str__(self):
        return self.stringify()

    def stringify(self, workbook, row=None, col=None):
        tokens = self._tokens[:]
        current_token = tokens.pop()
        if isinstance(current_token, NamePtg):
            return current_token.stringify(tokens, workbook, row, col)
        else:
            return current_token.stringify(tokens, workbook)

    @classmethod
    def parse(cls, data):
        return cls(TokenReader(data))

    @property
    def tokens(self):
        return self._tokens
