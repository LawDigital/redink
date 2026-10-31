"""One accounting contract shared by every LogPath chart. Independent of licensing."""
from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
import re


ALIASES = {k.casefold(): v for k, v in {
    'BallooMergePart_Word': 'BalloonMergePart_Word',
    'RI_BallooMergePart_Click': 'BalloonMergePart_Word',
    'RI_BalloonMergePart_Click': 'BalloonMergePart_Word',
    'Browser_SentToOutlook': 'Browser_SendToOutlook',
    'Easteregg_Word': 'EasterEgg_Word',
}.items()}
INTERNAL = frozenset('tool_loader tool_describe context_expand context_compact resolve_capability_route report_progress report_inability attest_response skill_use memory_put memory_get memory_list memory_delete'.split())
# Older clients used these agent_ names for ordinary filesystem tools, not subagents.
WORKSPACE_TOOLS = frozenset('agent_workspace_'+v for v in 'list read write file_op save_session_file search find_files move_to copy_to rename bulk_rename file_details recent_files create_folder_structure trash inventory_report'.split())
SURFACES = ('Word_Agent','Outlook_AutoPilot','Outlook_LocalChatAgent','Outlook_Tooling')
DETAIL = re.compile('AgentToolCall_('+'|'.join(SURFACES)+')_(.+)',re.I)
LEGACY = {'autopilot (local)', 'autopilot (mail)', 'autopilot (voicemail)'}


@dataclass(frozen=True)
class UsageLabel:
    kind: str
    label: str

    @property
    def primary(self):
        return self.kind in ('command', 'run')


@lru_cache(maxsize=4096)
def describe(feature: str) -> UsageLabel:
    name = ALIASES.get(feature.casefold(), feature)
    lower = name.casefold()
    if lower in {'agentrun_'+s.casefold() for s in SURFACES}:
        return UsageLabel('run', name)
    if lower in LEGACY:
        return UsageLabel('legacy', name)
    if lower.startswith('agenttoolcall_'):
        m = DETAIL.fullmatch(name)
        if not m:
            return UsageLabel('unclassified', 'Unrecognized tool-call surface')
        tool = m[2]
        low = tool.casefold()
        if low in INTERNAL:
            return UsageLabel('internal', tool)
        if low.startswith('knowledge_search_store_'):
            return UsageLabel('tool','Knowledge search · store name unavailable')
        if low.startswith('knowledge_search_store (') and tool.endswith(')'):
            return UsageLabel('tool','Knowledge search · '+tool[len('knowledge_search_store ('):-1])
        kind = ('skill' if low.startswith('skill_') else 'agent' if low.startswith('agent_') and low not in WORKSPACE_TOOLS else 'tool')
        return UsageLabel(kind,tool)
    return UsageLabel('command',name)


def primary_family(item, taxonomy):
    from .log_taxonomy import classify
    if item.kind == 'run':
        return {'id':'run-start','label':'Agent run starts','agentic':True}
    group = classify(item.label,taxonomy)
    if group['id'] == 'agent' and group['label'] == 'Agent-labelled operations':
        group = dict(group,label='Agent commands and controls')
    return dict(group,agentic=False)  # UI/control invocations are never evidence of a run.
