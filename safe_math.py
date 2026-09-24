"""Conservative plain-text math parser. Unsupported or ambiguous input is rejected."""
import re
from copy import deepcopy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def run(text):
    node = OxmlElement('m:r')
    props = OxmlElement('m:rPr')
    style = OxmlElement('m:sty')
    style.set(qn('m:val'), 'i' if re.fullmatch('[A-Za-z]', text) else 'p')
    props.append(style)
    node.append(props)
    value = OxmlElement('m:t')
    value.text = text
    node.append(value)
    return node


def parse(expression):
    # Never interpret adjacent letters as multiplication, or guess slash scope.
    if re.search(r'[A-Za-z]{2}|/.*[/]|/\s*\d+\s*[A-Za-z(]', expression):
        raise ValueError('Ambiguous letters or fraction scope')
    tokens = re.findall(r'\d+(?:\.\d+)?|[A-Za-z]|[²³]|[=+−\-×*÷/()√^]', expression)
    if ''.join(tokens) != re.sub(r'\s+', '', expression):
        raise ValueError('Unsupported mathematical notation')
    pos = 0

    def atom():
        nonlocal pos
        if pos >= len(tokens):
            raise ValueError('Missing operand')
        token = tokens[pos]
        pos += 1
        if token == '(':
            content = expr()
            if pos >= len(tokens) or tokens[pos] != ')':
                raise ValueError('Unclosed parenthesis')
            pos += 1
            nodes = [run('('), *content, run(')')]
        elif token == '√':
            if pos >= len(tokens):
                raise ValueError('Missing radicand')
            if tokens[pos] != '(' and not re.fullmatch(r'\d+(?:\.\d+)?', tokens[pos]):
                raise ValueError('Put a variable radicand in parentheses')
            grouped = tokens[pos] == '('
            if grouped:
                pos += 1
                content = expr()
                if pos >= len(tokens) or tokens[pos] != ')':
                    raise ValueError('Unclosed radicand')
                pos += 1
            else:
                content = [run(tokens[pos])]
                pos += 1
                if pos < len(tokens) and tokens[pos] in ('²', '³', '^'):
                    raise ValueError('Put a powered radicand in parentheses')
            rad = OxmlElement('m:rad')
            props = OxmlElement('m:radPr')
            hidden = OxmlElement('m:degHide'); hidden.set(qn('m:val'), '1')
            props.append(hidden)
            rad.append(props); rad.append(OxmlElement('m:deg'))
            body = OxmlElement('m:e'); body.extend(content); rad.append(body)
            nodes = [rad]
        elif re.fullmatch(r'\d+(?:\.\d+)?|[A-Za-z]', token):
            nodes = [run(token)]
        else:
            raise ValueError('Unexpected operand')
        if pos < len(tokens) and tokens[pos] in ('²', '³', '^'):
            exponent = tokens[pos]; pos += 1
            if exponent == '^':
                if pos >= len(tokens) or not re.fullmatch(r'\d+', tokens[pos]):
                    raise ValueError('Only explicit integer exponents are supported')
                exponent = tokens[pos]; pos += 1
            else:
                exponent = {'²': '2', '³': '3'}[exponent]
            power = OxmlElement('m:sSup')
            base = OxmlElement('m:e'); base.extend(nodes)
            sup = OxmlElement('m:sup'); sup.append(run(exponent))
            power.extend((base, sup)); nodes = [power]
        return nodes

    def term():
        nonlocal pos
        nodes = atom()
        while pos < len(tokens):
            token = tokens[pos]
            if token in ('*', '×', '÷', '/'):
                pos += 1
                right = atom()
                if token == '/':
                    frac = OxmlElement('m:f')
                    num = OxmlElement('m:num'); num.extend(nodes)
                    den = OxmlElement('m:den'); den.extend(right)
                    frac.extend((num, den)); nodes = [frac]
                else:
                    nodes += [run(token), *right]
            elif token == '(' or token == '√' or re.fullmatch('[A-Za-z]', token):
                nodes += atom()
            else:
                break
        return nodes

    def expr():
        nonlocal pos
        nodes = []
        if pos < len(tokens) and tokens[pos] in ('-', '−', '+'):
            nodes.append(run(tokens[pos])); pos += 1
        nodes += term()
        while pos < len(tokens) and tokens[pos] in ('+', '-', '−', '='):
            operator = tokens[pos]; pos += 1
            nodes += [run(operator), *term()]
        return nodes

    nodes = expr()
    if pos != len(tokens):
        raise ValueError('Unconsumed mathematical notation')
    equation = OxmlElement('m:oMath'); equation.extend(nodes)
    return equation


def replace_span(paragraph, start, end, equation):
    """Surgically replace a text span in ordinary Word runs, preserving outside styling."""
    offset = 0
    inserted = False
    for source in list(paragraph.runs):
        text = source.text
        left, right = max(0, start - offset), min(len(text), end - offset)
        offset += len(text)
        if left >= right:
            continue
        if any(n.tag not in {qn('w:rPr'), qn('w:t')} for n in source._r):
            raise ValueError('Complex Word run requires author review')
        for part, is_math in [(text[:left], False), (None, True), (text[right:], False)]:
            if is_math:
                if not inserted:
                    source._r.addprevious(equation); inserted = True
            elif part:
                node = OxmlElement('w:r')
                if source._r.rPr is not None:
                    node.append(deepcopy(source._r.rPr))
                value = OxmlElement('w:t'); value.text = part
                value.set(qn('xml:space'), 'preserve'); node.append(value)
                source._r.addprevious(node)
        source._r.getparent().remove(source._r)
