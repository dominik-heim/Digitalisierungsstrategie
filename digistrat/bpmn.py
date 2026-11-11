"""Export der Ist-Prozessaufnahme als BPMN-2.0-XML (inkl. Diagramm-Layout mit Swimlanes).

Die Dateien lassen sich in bpmn.io, Camunda Modeler, Signavio oder ADONIS importieren und dort
weiter verfeinern.
"""
from __future__ import annotations

import re
from pathlib import Path
from xml.sax.saxutils import escape

LANE_H = 140
COL_W = 170
TASK_W, TASK_H = 120, 70


def _id(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9_]", "_", s)


def export_bpmn(name: str, proc: dict, path: Path) -> Path:
    from .report_pptx import _levels  # gleiche Spaltenlogik wie in der Folie

    steps, lanes, succ = proc["steps"], proc["lanes"], proc["succ"]
    lv = _levels(steps, succ)
    pid = "Process_" + _id(name)
    elems, flows, shapes, edges, lane_xml = [], [], [], [], []
    pos = {}
    slot: dict[tuple[int, int], int] = {}
    x0 = 60
    for st in steps:
        nr = st["nr"]
        li = lanes.index(st.get("lane") or "–")
        col = lv[nr]
        k = slot.get((li, col), 0)
        slot[(li, col)] = k + 1
        typ = st.get("typ") or "Aufgabe"
        tag = {"Start": "startEvent", "Ende": "endEvent", "Entscheidung": "exclusiveGateway"}.get(typ, "task")
        w, h = (36, 36) if tag.endswith("Event") else (50, 50) if tag == "exclusiveGateway" else (TASK_W, TASK_H)
        cx = x0 + 120 + col * COL_W + COL_W / 2
        cy = li * LANE_H + LANE_H / 2 + k * 20
        eid = f"N{nr}"
        pos[nr] = (cx, cy, w, h)
        doc = escape(" | ".join(x for x in [st.get("system"), st.get("medium"), st.get("bemerkung")] if x))
        elems.append(f'    <bpmn:{tag} id="{eid}" name="{escape(st["schritt"])}">'
                     + (f"<bpmn:documentation>{doc}</bpmn:documentation>" if doc else "") + f"</bpmn:{tag}>")
        shapes.append(f'      <bpmndi:BPMNShape id="{eid}_di" bpmnElement="{eid}">'
                      f'<dc:Bounds x="{cx - w / 2:.0f}" y="{cy - h / 2:.0f}" width="{w}" height="{h}" />'
                      f'</bpmndi:BPMNShape>')
    breaks = {(b["von_nr"], b["nach_nr"]) for b in proc["medienbrueche"]}
    for a, bs in succ.items():
        for b in bs:
            if a not in pos or b not in pos:
                continue
            fid = f"F{a}_{b}"
            label = ' name="Medienbruch"' if (a, b) in breaks else ""
            flows.append(f'    <bpmn:sequenceFlow id="{fid}" sourceRef="N{a}" targetRef="N{b}"{label} />')
            ax, ay, aw, _ = pos[a]
            bx, by, bw, _ = pos[b]
            pts = [(ax + aw / 2, ay), (bx - bw / 2, by)]
            if abs(ay - by) > 1 and bx > ax:
                xm = (ax + aw / 2 + bx - bw / 2) / 2
                pts = [(ax + aw / 2, ay), (xm, ay), (xm, by), (bx - bw / 2, by)]
            wp = "".join(f'<di:waypoint x="{x:.0f}" y="{y:.0f}" />' for x, y in pts)
            edges.append(f'      <bpmndi:BPMNEdge id="{fid}_di" bpmnElement="{fid}">{wp}</bpmndi:BPMNEdge>')
    ncol = max(lv.values(), default=0) + 1
    width = 120 + ncol * COL_W + 40
    for i, ln in enumerate(lanes):
        lid = f"Lane_{i}"
        refs = "".join(f"<bpmn:flowNodeRef>N{s['nr']}</bpmn:flowNodeRef>" for s in steps
                       if (s.get("lane") or "–") == ln)
        lane_xml.append(f'      <bpmn:lane id="{lid}" name="{escape(ln)}">{refs}</bpmn:lane>')
        shapes.insert(0, f'      <bpmndi:BPMNShape id="{lid}_di" bpmnElement="{lid}" isHorizontal="true">'
                         f'<dc:Bounds x="{x0 + 30}" y="{i * LANE_H}" width="{width - 30}" height="{LANE_H}" />'
                         f'</bpmndi:BPMNShape>')
    shapes.insert(0, f'      <bpmndi:BPMNShape id="Participant_di" bpmnElement="Participant" isHorizontal="true">'
                     f'<dc:Bounds x="{x0}" y="0" width="{width}" height="{len(lanes) * LANE_H}" /></bpmndi:BPMNShape>')
    xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<bpmn:definitions xmlns:bpmn="http://www.omg.org/spec/BPMN/20100524/MODEL"
  xmlns:bpmndi="http://www.omg.org/spec/BPMN/20100524/DI" xmlns:dc="http://www.omg.org/spec/DD/20100524/DC"
  xmlns:di="http://www.omg.org/spec/DD/20100524/DI" id="Defs_{_id(name)}" targetNamespace="http://digistrat/bpmn"
  exporter="digistrat" exporterVersion="1.0">
  <bpmn:collaboration id="Collab">
    <bpmn:participant id="Participant" name="{escape(name)}" processRef="{pid}" />
  </bpmn:collaboration>
  <bpmn:process id="{pid}" name="{escape(name)}" isExecutable="false">
    <bpmn:laneSet id="LaneSet">
{chr(10).join(lane_xml)}
    </bpmn:laneSet>
{chr(10).join(elems)}
{chr(10).join(flows)}
  </bpmn:process>
  <bpmndi:BPMNDiagram id="Diagram">
    <bpmndi:BPMNPlane id="Plane" bpmnElement="Collab">
{chr(10).join(shapes)}
{chr(10).join(edges)}
    </bpmndi:BPMNPlane>
  </bpmndi:BPMNDiagram>
</bpmn:definitions>
'''
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(xml, encoding="utf-8")
    return path
