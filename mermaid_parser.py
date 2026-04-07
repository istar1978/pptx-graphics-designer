#!/usr/bin/env python3
"""
Mermaid Parser for PPTX Graphics Designer
Parses Mermaid syntax and converts to PPT diagram elements
"""

import re
from enum import Enum

class DiagramType(Enum):
    FLOWCHART = "flowchart"
    SEQUENCE = "sequence"
    GANTT = "gantt"
    CLASS = "class"
    STATE = "state"
    ER = "er"
    PIE = "pie"

class MermaidParser:
    """Parser for Mermaid diagram syntax"""

    def __init__(self):
        self.nodes = {}
        self.edges = []
        self.diagram_type = None

    def parse(self, mermaid_code):
        """
        Parse Mermaid code and return diagram structure

        Args:
            mermaid_code: String containing Mermaid syntax

        Returns:
            dict: Parsed diagram structure
        """
        # Support single-line semicolon-separated Mermaid input by converting to lines
        if isinstance(mermaid_code, str) and ';' in mermaid_code and '\n' not in mermaid_code:
            lines = [part.strip() for part in mermaid_code.split(';') if part.strip()]
        else:
            lines = [line.strip() for line in mermaid_code.split('\n') if line.strip()]

        if not lines:
            raise ValueError("Empty Mermaid code")

        # Detect diagram type
        first_line = lines[0].lower()
        if first_line.startswith('graph'):
            self.diagram_type = DiagramType.FLOWCHART
            return self._parse_flowchart(lines)
        elif first_line.startswith('sequence'):
            self.diagram_type = DiagramType.SEQUENCE
            return self._parse_sequence(lines)
        elif first_line.startswith('gantt'):
            self.diagram_type = DiagramType.GANTT
            return self._parse_gantt(lines)
        elif first_line.startswith('class'):
            self.diagram_type = DiagramType.CLASS
            return self._parse_class(lines)
        elif first_line.startswith('state'):
            self.diagram_type = DiagramType.STATE
            return self._parse_state(lines)
        elif first_line.startswith('er'):
            self.diagram_type = DiagramType.ER
            return self._parse_er(lines)
        else:
            # Default to flowchart
            self.diagram_type = DiagramType.FLOWCHART
            return self._parse_flowchart(lines)

    def _parse_flowchart(self, lines):
        """Parse flowchart diagram"""
        nodes = {}
        edges = []

        for line in lines:
            # Match edges like: A --> B, A --- B, A -->|label| B
            m = re.match(r'^(.*?)\s*(?:-->|---)\s*(?:\|([^|]+)\|\s*)?(.*)$', line)
            if m:
                source = m.group(1).strip()
                label = (m.group(2) or '').strip()
                target = m.group(3).strip()

                source_node = self._parse_node(source)
                target_node = self._parse_node(target)

                if source_node:
                    nodes[source_node['id']] = source_node
                if target_node:
                    nodes[target_node['id']] = target_node

                edges.append({
                    'source': source_node['id'] if source_node else source,
                    'target': target_node['id'] if target_node else target,
                    'label': label,
                    'style': 'solid' if '-->' in line else 'dashed'
                })
            else:
                # Node definition
                node = self._parse_node(line)
                if node:
                    nodes[node['id']] = node

        return {
            'type': 'flowchart',
            'nodes': list(nodes.values()),
            'edges': edges
        }

    def _parse_node(self, node_str):
        """Parse individual node definition"""
        node_str = node_str.strip()

        # Handle different node formats
        # A[Label]
        match = re.match(r'^(\w+)\[([^\]]+)\]$', node_str)
        if match:
            return {
                'id': match.group(1),
                'label': match.group(2),
                'shape': 'rectangle'
            }

        # A(Label)
        match = re.match(r'^(\w+)\(([^)]+)\)$', node_str)
        if match:
            return {
                'id': match.group(1),
                'label': match.group(2),
                'shape': 'roundrectangle'
            }

        # A{Label}
        match = re.match(r'^(\w+)\{([^{}]+)\}$', node_str)
        if match:
            return {
                'id': match.group(1),
                'label': match.group(2),
                'shape': 'diamond'
            }

        # A>Label]
        match = re.match(r'^(\w+)> ([^\]]+)\]$', node_str)
        if match:
            return {
                'id': match.group(1),
                'label': match.group(2),
                'shape': 'parallelogram'
            }

        # Simple node
        if re.match(r'^\w+$', node_str):
            return {
                'id': node_str,
                'label': node_str,
                'shape': 'rectangle'
            }

        return None

    def _parse_sequence(self, lines):
        """Parse sequence diagram"""
        participants = []
        messages = []

        for line in lines[1:]:  # Skip 'sequenceDiagram'
            line = line.strip()
            if line.startswith('participant'):
                parts = line.split()
                if len(parts) >= 2:
                    participants.append({
                        'id': parts[1],
                        'label': ' '.join(parts[2:]) if len(parts) > 2 else parts[1]
                    })
            elif '->>' in line or '->' in line:
                parts = re.split(r'\s*->>?\s*', line)
                if len(parts) >= 2:
                    source = parts[0].strip()
                    rest = parts[1].split(':', 1)
                    target = rest[0].strip()
                    message = rest[1].strip() if len(rest) > 1 else ""

                    messages.append({
                        'source': source,
                        'target': target,
                        'message': message,
                        'type': 'sync' if '->>' in line else 'async'
                    })

        return {
            'type': 'sequence',
            'participants': participants,
            'messages': messages
        }

    def _parse_gantt(self, lines):
        """Parse Gantt chart"""
        tasks = []

        for line in lines[1:]:  # Skip 'gantt'
            line = line.strip()
            if ':' in line:
                parts = line.split(':', 1)
                task_name = parts[0].strip()
                task_def = parts[1].strip()

                # Simple parsing - can be enhanced
                tasks.append({
                    'name': task_name,
                    'definition': task_def
                })

        return {
            'type': 'gantt',
            'tasks': tasks
        }

    def _parse_class(self, lines):
        """Parse class diagram"""
        classes = []
        relationships = []

        current_class = None
        for line in lines[1:]:  # Skip 'classDiagram'
            line = line.strip()
            if line.startswith('class '):
                parts = line.split()
                class_name = parts[1]
                current_class = {
                    'name': class_name,
                    'attributes': [],
                    'methods': []
                }
                classes.append(current_class)
            elif current_class and line.startswith('+') or line.startswith('-'):
                if '(' in line:
                    current_class['methods'].append(line)
                else:
                    current_class['attributes'].append(line)
            elif '<|--' in line or '*--' in line or 'o--' in line:
                # Relationship
                parts = re.split(r'\s*(<\|--|\*--|o--|-->|--)\s*', line)
                if len(parts) >= 3:
                    relationships.append({
                        'source': parts[0],
                        'target': parts[2],
                        'type': parts[1]
                    })

        return {
            'type': 'class',
            'classes': classes,
            'relationships': relationships
        }

    def _parse_state(self, lines):
        """Parse state diagram"""
        states = []
        transitions = []

        for line in lines[1:]:  # Skip 'stateDiagram'
            line = line.strip()
            if '-->' in line:
                parts = line.split('-->', 1)
                if len(parts) == 2:
                    source = parts[0].strip()
                    rest = parts[1].split(':', 1)
                    target = rest[0].strip()
                    label = rest[1].strip() if len(rest) > 1 else ""

                    transitions.append({
                        'source': source,
                        'target': target,
                        'label': label
                    })

                    # Add states if not already present
                    for state_id in [source, target]:
                        if not any(s['id'] == state_id for s in states):
                            states.append({
                                'id': state_id,
                                'label': state_id
                            })

        return {
            'type': 'state',
            'states': states,
            'transitions': transitions
        }

    def _parse_er(self, lines):
        """Parse ER diagram"""
        entities = []
        relationships = []

        for line in lines[1:]:  # Skip 'erDiagram'
            line = line.strip()
            if '||--o{' in line or '||--||' in line or '}o--||' in line:
                # Relationship
                parts = re.split(r'\s*(\|\|--o\{|\|\|--\|\||\}o--\|\|)\s*', line)
                if len(parts) >= 3:
                    relationships.append({
                        'entity1': parts[0],
                        'entity2': parts[2],
                        'type': parts[1]
                    })
            elif line and not line.startswith('erDiagram'):
                # Entity definition
                entities.append({'name': line})

        return {
            'type': 'er',
            'entities': entities,
            'relationships': relationships
        }

def parse_mermaid(mermaid_code):
    """Convenience function to parse Mermaid code"""
    parser = MermaidParser()
    return parser.parse(mermaid_code)