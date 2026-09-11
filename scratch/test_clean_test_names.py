import glob
import json
import re

KNOWN_EXACT_TEST_NAMES = [
    r'Breast\s+and\s+Ovarian\s+(?:Cancer\s+)?Extended\s+Panel\s*[-–]\s*Liquid\s+Biopsy\s+Assay',
    r'(?:Liquidseq\s+Actionable|Brainseq)\s+Genomic\s+Profiling\s+Panel(?:\s*[-–]\s*Advance)?',
    r'Liquidseq\s+Comprehensive\s+Genomic\s+Profile\s*\([A-Z]+\)\s*Panel',
    r'Liquidseq\s+Lung\s+Cancer\s+Panel',
    r'Solidseq\s+Comprehensive\s+Panel(?:\s+On\s+(?:the\s+)?Illumina\s+[\w\s-]+\s+Platform)?',
    r'Whole\s+Exome\s+Sequencing(?:\s+on\s+(?:the\s+)?Illumina\s+[\w\s-]+\s+Platform)?',
]

test_name_pattern = re.compile(
    r'(?:\bTEST\s+NAME\b(?!\s*[:=])|' + '|'.join(KNOWN_EXACT_TEST_NAMES) + r')',
    re.IGNORECASE
)

for f in glob.glob('extracted_jsons/*.json'):
    with open(f, 'r', encoding='utf-8') as fp:
        d = json.load(fp)
    
    found_matches = []
    def scan_for_matches(obj):
        if isinstance(obj, str):
            for m in test_name_pattern.finditer(obj):
                found_matches.append(m.group(0))
        elif isinstance(obj, dict):
            for v in obj.values():
                scan_for_matches(v)
        elif isinstance(obj, list):
            for item in obj:
                scan_for_matches(item)
                
    scan_for_matches(d)
    unique_matches = list(set(found_matches))
    print(f"{f.split('/')[-1].split(chr(92))[-1][:35]}: {len(found_matches)} matches -> {unique_matches}")
