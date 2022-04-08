"""
Participant Parser v1.0 BETA

An extension program for MSDACSYS Participants

(c) 2022-present Verdadero, Ycong
"""

## Needs improvement
##
## • Take note of the two-line (2) participant that should be included
## • Ignore the second name after the first tag with "@" — Assuming that the first name entry is already the participant
## • Ignore the parentheses in the first name
## • When converting to name case, ignore the non-alphabet capitalization such as double-quotes and make the 2nd character the capital instead.

import re, sys

class ParticipantParser(object):
    def __init__(self):
        self.NL = '\n'                                                                                              ## Newline
        self.DELIM = ':'                                                                                            ## Search delimiter
        self.FILE_READ = 'message.txt'                                                                              ## Message Source
        self.FILTER_EXCLUSIONS = ['AY', 'to'], ['TBA']                                                              ## Exclusion for string formatting (Role, Name)


    def getMessage(self):
        """
        Returns a string of participants from a message file
        """
        with open(self.FILE_READ, 'r', encoding='utf-8') as f:
            return f.read()

    
    def fmtStr(self, string:str, mode):
        """
        Formats ROLE and NAME to a presentable format
        """
        out = []
        for word in string.split():
            if not mode:                                                                                            ## ROLE Formatting
                if word not in self.FILTER_EXCLUSIONS[0]: out.append(f"{word[:1].upper()}{word[1:].lower()}")
                else: out.append(word)

            else:                                                                                                   ## NAME Formatting
                word = word.replace('@', '')
                if word not in self.FILTER_EXCLUSIONS[1]: out.append(f"{word[:1].upper()}{word[1:].lower()}")
                else: out.append(word)

        return ' '.join(out)
    

    def parse(self, source=None):
        """
        Returns a tuple of participants with its ROLE and NAME
        """
        RESULT = []
        SOURCE = self.getMessage() if source is None else source

        ## Strip lines
        CLEAN = '\n'.join([line.strip() for line in SOURCE.splitlines()])                                           ## Strips the spaces for both ends of the string

        ## Distinguish roles and names
        for i, line in enumerate(CLEAN.splitlines()):
            if self.DELIM in line:
                PART = line.split(self.DELIM)
                ROLE, NAME = PART[0].strip(), PART[1].strip()

                if len(re.sub(r"[^A-Za-z]+", '', NAME)):                                                            ## If there's a name in the line
                    RESULT.append((self.fmtStr(ROLE, 0), self.fmtStr(NAME, 1)))

                else:                                                                                               ## When there's no NAME next to the ROLE
                    if i+1 == len(CLEAN.splitlines()): NLPT = ''
                    else: NLPT = CLEAN.split(self.NL)[i+1].strip().replace('@', '')                                 ## Next Line Participant

                    if self.DELIM in NLPT: NLPT = ''                                                                ## Ignore NAME when it's a ROLE
                    RESULT.append((self.fmtStr(line.replace(self.DELIM, ''), 0), self.fmtStr(NLPT.strip(), 1)))

        return RESULT
    
    
    def showResult(self, source=None, out=sys.stdout):
        for i in self.parse() if source is None else source:
            print(f"{i[0]}: {i[1]}", file=out)



if __name__ == '__main__':
    PARSER = ParticipantParser()
    PARSER.parse()
    # PARSER.showResult()
