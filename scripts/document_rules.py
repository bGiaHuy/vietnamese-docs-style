"""Shared, explicit formatting rules. Defaults are not compliance certificates."""
import json
import math
from copy import deepcopy
from pathlib import Path

PROFILES = ('office', 'administrative', 'academic', 'proposal', 'minutes-administrative',
            'minutes-general', 'party', 'custom')
# Profiles with an official layout: hidden first-page PAGE field, 1 cm first-line
# indent and plain-text dash bullets. 'party' follows Hướng dẫn 05-HD/VPTW (2026).
ADMIN = ('administrative', 'minutes-administrative', 'party')
PARTY_MARGINS_CM = {'left': [3.0, 3.0], 'right': [1.5, 1.5], 'top': [2.0, 2.0], 'bottom': [2.0, 2.0]}


def load_rules(profile='administrative', config=None):
    if profile not in PROFILES:
        raise ValueError(f'Unknown profile: {profile}')
    rules = {
        'profile': profile, 'source': 'Skill defaults; not a compliance certificate',
        'font': None if profile == 'custom' else 'Times New Roman',
        'color': None if profile == 'custom' else '000000',
        'body_size_pt': 14 if profile == 'party' else 13, 'line_spacing': 1.15,
        'line_spacing_exact_pt': 20 if profile == 'party' else None,
        'page_number_size_pt': 14 if profile == 'party' else 13,
        'page_cm': None if profile == 'custom' else [21.0, 29.7],
        'allow_landscape': profile != 'custom',
        'margins_cm': (deepcopy(PARTY_MARGINS_CM) if profile == 'party'
                       else {'left': [3.0, 3.5], 'right': [1.5, 2.0],
                             'top': [2.0, 2.5], 'bottom': [2.0, 2.5]}
                       if profile in ADMIN else {}),
        'sections': [], 'style_rules': {}, 'required_text': [], 'forbid_bullets': False,
        'forbid_shading': False,
    }
    if config is not None:
        overrides = (json.loads(Path(config).read_text(encoding='utf-8-sig'))
                     if isinstance(config, (str, Path)) else deepcopy(config))
        if not isinstance(overrides, dict):
            raise ValueError('Rules must be a JSON object')
        unknown = set(overrides) - set(rules)
        if unknown:
            raise ValueError(f'Unknown rules: {sorted(unknown)}')
        if 'profile' in overrides and overrides['profile'] != profile:
            raise ValueError('Config profile differs from selected profile')
        rules.update(overrides)
    _validate(rules)
    return rules


def _validate(rules):
    def number(value):
        return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)

    def page(value):
        if value is not None and (not isinstance(value, list) or len(value) != 2
                                  or any(not number(x) or x <= 0 for x in value)):
            raise ValueError('page_cm must be null or two positive numbers')

    def margins(value):
        if not isinstance(value, dict) or set(value) - {'left', 'right', 'top', 'bottom'}:
            raise ValueError('Invalid margins_cm')
        for bounds in value.values():
            if (not isinstance(bounds, list) or len(bounds) != 2
                    or any(not number(x) or x < 0 for x in bounds)
                    or bounds[0] > bounds[1]):
                raise ValueError('Margins need [minimum, maximum] in cm')

    page(rules['page_cm'])
    margins(rules['margins_cm'])
    if not isinstance(rules['sections'], list):
        raise ValueError('sections must be a list')
    for section in rules['sections']:
        if not isinstance(section, dict) or set(section) - {'page_cm', 'margins_cm', 'allow_landscape'}:
            raise ValueError('Invalid per-section rules')
        if 'page_cm' in section:
            page(section['page_cm'])
        if 'margins_cm' in section:
            margins(section['margins_cm'])
        if 'allow_landscape' in section and not isinstance(section['allow_landscape'], bool):
            raise ValueError('allow_landscape must be boolean')
    for key in ('allow_landscape', 'forbid_bullets', 'forbid_shading'):
        if not isinstance(rules[key], bool):
            raise ValueError(f'{key} must be boolean')
    if rules['font'] is not None and not isinstance(rules['font'], str):
        raise ValueError('font must be a string or null')
    color = rules['color']
    if color is not None:
        if not isinstance(color, str) or len(color) != 6 or any(c not in '0123456789abcdefABCDEF' for c in color):
            raise ValueError('color must be six hexadecimal digits or null')
        rules['color'] = color.upper()
    for key in ('body_size_pt', 'line_spacing', 'page_number_size_pt'):
        if not number(rules[key]) or rules[key] <= 0:
            raise ValueError(f'{key} must be positive')
    exact = rules['line_spacing_exact_pt']
    if exact is not None and (not number(exact) or exact <= 0):
        raise ValueError('line_spacing_exact_pt must be null or positive')
    if (not isinstance(rules['required_text'], list)
            or any(not isinstance(x, str) or not x for x in rules['required_text'])):
        raise ValueError('required_text must contain nonempty literal strings')
    if not isinstance(rules['source'], str):
        raise ValueError('source must be a string')
    if not isinstance(rules['style_rules'], dict):
        raise ValueError('style_rules must be an object keyed by style name')
    for name, spec in rules['style_rules'].items():
        if not isinstance(name, str) or not name or not isinstance(spec, dict):
            raise ValueError('Invalid named style rule')
        if set(spec) - {'font', 'color', 'size_pt', 'bold', 'italic'}:
            raise ValueError('Unknown named style property')
        for key in ('bold', 'italic'):
            if key in spec and not isinstance(spec[key], bool):
                raise ValueError(f'{key} must be boolean')
        if 'size_pt' in spec:
            bounds = spec['size_pt']
            if (not isinstance(bounds, list) or len(bounds) != 2
                    or any(not number(x) or x <= 0 for x in bounds) or bounds[0] > bounds[1]):
                raise ValueError('size_pt needs [minimum, maximum]')
        if 'font' in spec and (not isinstance(spec['font'], str) or not spec['font']):
            raise ValueError('Named style font must be nonempty')
        if 'color' in spec:
            color = spec['color']
            if (not isinstance(color, str) or len(color) != 6
                    or any(c not in '0123456789abcdefABCDEF' for c in color)):
                raise ValueError('Invalid named style color')
            spec['color'] = color.upper()


def section_rules(rules, index):
    result = dict(rules)
    if index < len(rules['sections']):
        result.update(rules['sections'][index])
    return result
