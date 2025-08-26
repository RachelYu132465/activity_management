import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.actions.influencer import iter_dicts, build_profile, build_people

DATA_DIR = Path('data/shared')


def test_iter_dicts_flattens_nested_lists():
    nested = [{'name': 'A'}, [{'name': 'B'}, {'name': 'C'}]]
    names = {d['name'] for d in iter_dicts(nested)}
    assert names == {'A', 'B', 'C'}


def test_build_profile_includes_org_education_and_experience():
    info = {
        'current_position': {'organization': 'Org', 'title': 'T'},
        'highest_education': {'school': 'S', 'department': 'D'},
        'experience': ['E1', 'E2']
    }
    profile = build_profile(info)
    assert 'Org' in profile
    assert 'S D' in profile
    assert 'E1' in profile and 'E2' in profile


def test_build_people_matches_influencer_names():
    programs = json.loads((DATA_DIR / 'program_data.json').read_text(encoding='utf-8'))
    influencers = json.loads((DATA_DIR / 'influencer_data.json').read_text(encoding='utf-8'))
    first_event = programs[0]
    chairs, speakers = build_people(first_event, influencers)
    by_name = {p['name']: p for p in chairs + speakers}
    for name in ['林世嘉', '林奕汝', '中山功一', '官建村']:
        assert by_name[name]['profile'], name
