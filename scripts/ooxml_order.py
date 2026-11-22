"""ECMA-376 child-element ordering and canonical non-target preservation checks."""
import hashlib
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

LEADING_PROPERTY = {
    'p': 'pPr',
    'r': 'rPr',
    'tbl': 'tblPr',
    'tr': 'trPr',
    'tc': 'tcPr',
}

# Child sequence per ISO/IEC 29500-1 (ECMA-376) WordprocessingML complex types.
ORDERED_CONTAINERS = {
    'tblPr': (
        'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
        'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd',
        'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook',
        'tblCaption', 'tblDescription', 'tblPrChange',
    ),
    'trPr': (
        'cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter',
        'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden',
        'ins', 'del', 'trPrChange',
    ),
    'tcPr': (
        'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd',
        'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark',
        'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange',
    ),
    'pPr': (
        'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
        'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
        'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
        'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
        'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
        'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
        'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr',
        'pPrChange',
    ),
    'rPr': (
        'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
        'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
        'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w',
        'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr',
        'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
        'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange',
    ),
    'sectPr': (
        'headerReference', 'footerReference', 'footnotePr', 'endnotePr',
        'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
        'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
        'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings',
        'sectPrChange',
    ),
    'tblCellMar': ('top', 'start', 'left', 'bottom', 'end', 'right'),
    'tcMar': ('top', 'start', 'left', 'bottom', 'end', 'right'),
}


def _w_localname(element):
    if not isinstance(element.tag, str):
        return None
    qname = etree.QName(element)
    return qname.localname if qname.namespace == W else None


def normalize_ordered_children(parent):
    """Sort children of an ECMA-376 ordered container while preserving unknown tags."""
    parent_local = _w_localname(parent)
    if parent_local is None:
        return parent

    order = ORDERED_CONTAINERS.get(parent_local)
    lead = LEADING_PROPERTY.get(parent_local)
    children = list(parent)
    if not children:
        return parent

    if lead is not None:
        # Ensure leading property element sits at index 0 while leaving other children untouched.
        lead_nodes = [c for c in children if _w_localname(c) == lead]
        rest_nodes = [c for c in children if _w_localname(c) != lead]
        if lead_nodes:
            for c in children:
                parent.remove(c)
            for c in lead_nodes + rest_nodes:
                parent.append(c)
        return parent

    if order is None:
        return parent

    change_tag = f'{parent_local}Change'
    rank_map = {tag: idx for idx, tag in enumerate(order)}

    def _sort_key(elem_and_idx):
        elem, orig_idx = elem_and_idx
        c_local = _w_localname(elem)
        if c_local == change_tag:
            return (2, len(order), orig_idx)
        if c_local in rank_map:
            return (0, rank_map[c_local], orig_idx)
        return (1, 0, orig_idx)

    ordered = [elem for _, elem in sorted(enumerate(children), key=lambda pair: _sort_key((pair[1], pair[0])))]
    if [id(e) for e in ordered] != [id(e) for e in children]:
        for c in children:
            parent.remove(c)
        for c in ordered:
            parent.append(c)
    return parent


def insert_ordered_child(parent, child, replace_existing=True):
    """Insert or replace a WordprocessingML child at its valid schema slot."""
    parent_local = _w_localname(parent)
    child_local = _w_localname(child)
    if parent_local is None or child_local is None:
        raise ValueError('insert_ordered_child expects WordprocessingML elements')

    if replace_existing:
        for existing in list(parent):
            if _w_localname(existing) == child_local:
                parent.remove(existing)

    # Leading property container must always be child 0 when present.
    if LEADING_PROPERTY.get(parent_local) == child_local:
        parent.insert(0, child)
        return child

    order = ORDERED_CONTAINERS.get(parent_local)
    if order is None or child_local not in order:
        parent.append(child)
        return child

    normalize_ordered_children(parent)
    target_rank = order.index(child_local)
    for idx, sibling in enumerate(parent):
        sib_local = _w_localname(sibling)
        if sib_local in order and order.index(sib_local) > target_rank:
            parent.insert(idx, child)
            return child
    parent.append(child)
    return child


def check_element_order(root):
    """Verify leading-property and sibling sequence rules across all containers."""
    violations = []
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        local = _w_localname(elem)
        if local is None:
            continue

        lead = LEADING_PROPERTY.get(local)
        if lead is not None:
            elem_children = [c for c in elem if isinstance(c.tag, str)]
            for idx, child in enumerate(elem_children):
                if _w_localname(child) == lead and idx > 0:
                    violations.append({
                        'container': local,
                        'rule': 'leading_property',
                        'message': f'<w:{lead}> must be the first child of <w:{local}> (found at position {idx})',
                    })

        order = ORDERED_CONTAINERS.get(local)
        if order is not None:
            highest_rank = -1
            highest_tag = ''
            for child in elem:
                if not isinstance(child.tag, str):
                    continue
                c_local = _w_localname(child)
                if c_local not in order:
                    continue
                rank = order.index(c_local)
                if rank < highest_rank:
                    violations.append({
                        'container': local,
                        'rule': 'child_order',
                        'message': f'<w:{c_local}> must precede <w:{highest_tag}> inside <w:{local}>',
                    })
                else:
                    highest_rank = rank
                    highest_tag = c_local
    return violations


def canonical_node_sha256(element):
    """Compute canonical XML SHA-256 over a subtree for non-target diffing."""
    try:
        payload = etree.tostring(element, method='c14n2', strip_text=False)
    except (ValueError, etree.C14NError):
        try:
            payload = etree.tostring(element, method='c14n')
        except (ValueError, etree.C14NError):
            payload = etree.tostring(element, encoding='utf-8')
    return hashlib.sha256(payload).hexdigest()


def check_non_target_preservation(before_root, after_root, allowed_body_indices):
    """Verify all w:body direct children outside allowed_body_indices are unchanged."""
    before_body = before_root.find(f'{{{W}}}body')
    after_body = after_root.find(f'{{{W}}}body')
    if before_body is None or after_body is None:
        raise ValueError('Document XML is missing w:body')

    before_children = list(before_body)
    after_children = list(after_body)
    allowed = set(allowed_body_indices)

    if len(before_children) != len(after_children):
        return {
            'xml_ntpr': 0.0,
            'total_non_target': max(0, len(before_children) - len(allowed)),
            'preserved_non_target': 0,
            'unexpected_indices': [],
            'structure_changed': True,
        }

    total_non_target = 0
    preserved = 0
    unexpected = []
    for idx, (b_child, a_child) in enumerate(zip(before_children, after_children)):
        if idx in allowed:
            continue
        total_non_target += 1
        if canonical_node_sha256(b_child) == canonical_node_sha256(a_child):
            preserved += 1
        else:
            unexpected.append({
                'body_index': idx,
                'tag': _w_localname(a_child) or str(a_child.tag),
            })

    ntpr = 1.0 if total_non_target == 0 else preserved / total_non_target
    return {
        'xml_ntpr': round(ntpr, 6),
        'total_non_target': total_non_target,
        'preserved_non_target': preserved,
        'unexpected_indices': unexpected,
        'structure_changed': False,
    }
