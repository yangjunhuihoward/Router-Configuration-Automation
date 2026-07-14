# Nokia SR-OS Automated Configuration Builder

A lightweight, zero-hardcode network configuration automation engine for Nokia SR-OS routers.  
Built by a network engineer, for network engineers — **no Python knowledge required to use**.

---

## ✨ What It Does

Given an **Excel workbook** with device data and a set of **Jinja2 templates** defining SR-OS configuration structure, this tool automatically generates one `.cfg` file per router in seconds.

```
Excel Workbook  +  Jinja2 Templates  →  Python Engine  →  R1.cfg / R2.cfg / R3.cfg ...
```

---

## 🎯 Design Philosophy

| Principle | Detail |
|-----------|--------|
| **Zero hardcoding** | The Python engine contains no sheet names, column names, or service-specific logic |
| **Excel = Source of truth** | All device data lives in Excel — one workbook per project |
| **Column name = Variable name** | Whatever you name a column in Excel is the exact variable name in Jinja2 |
| **Engine never changes** | Only the Excel file and Jinja2 templates change between projects |
| **Engineer friendly** | If you know Excel and SR-OS CLI, you can use this tool |

---

## 🗂️ Project Structure

```
project/
├── config_builder_pydantic_dict.py   ← Python engine (never modify)
├── Spreadsheet/
│   └── Spreadsheet v1.xlsx           ← Your Excel workbook
├── templates/
│   ├── 7250-IXR-ALL.j2               ← Master template (includes sub-templates)
│   └── 7250-IXR-R6D/
│       ├── 1_device_7250r6d_system.j2
│       ├── 8_device_7250r6d_Router_IGP.j2
│       └── 9_device_7250r6d_iBGP.j2
└── Configuration-Files/              ← Generated .cfg files (auto-created)
    ├── R1.cfg
    ├── R2.cfg
    └── ...
```

---

## 🚀 Quick Start

### 1. Install dependencies

```bash
pip install pandas openpyxl jinja2 pydantic
```

### 2. Prepare your Excel workbook

- Create a sheet named **`Node`** with at least `NodeName` and `TemplateName` columns
- Add service sheets (e.g. `P2P`, `VPRN1000`, `VPLS2600`) with a `NodeName` column
- Optionally add global sheets (e.g. `SERVER`, `QOS`) **without** a `NodeName` column

### 3. Write your Jinja2 templates

```jinja2
configure
    system
        name "{{ node.NodeName }}"
    exit
    router Base
        interface "system"
            address {{ node.SystemIP }}/32
        exit
    exit
```

### 4. Run the engine

```bash
python config_builder_pydantic_dict.py
```

```
⚡ [Automation Engine] Parsing original Excel sheets automatically...
  Global sheets loaded: ['SERVER', 'QOS']
⚡ [Automation Engine] Starting to render configurations for 40 devices...
  -> Generated successfully: R1.cfg associated with template: '7250-IXR-ALL.j2'
  -> Generated successfully: R2.cfg associated with template: '7250-IXR-ALL.j2'
  ...
=======================================================
🎉 Finished!  Success: 40  Failed: 0
```

---

## 📊 Excel Workbook Design

### Sheet Types

| Sheet Type | Rule | Example Sheets |
|------------|------|----------------|
| **Node sheet** | Must be named `Node`. Contains one row per device. | `Node` |
| **Device sheets** | Must have a `NodeName` column. One or multiple rows per device. | `P2P`, `VPRN1000`, `VPLS2600`, `MDA` |
| **Global sheets** | Must NOT have a `NodeName` column. Shared across all devices. | `SERVER`, `QOS`, `POLICY` |

### Node Sheet (Required)

| Column | Required | Description |
|--------|----------|-------------|
| `NodeName` | ✅ Yes | Unique device name. Used as key across all sheets |
| `TemplateName` | ✅ Yes | Jinja2 template filename e.g. `7250-IXR-ALL.j2` |
| `SystemIP` | Recommended | Device loopback/system IP |
| `NodeType` | Recommended | Platform type e.g. `7250-IXR-R6D`, `7705-SAR8` |
| Any other column | Optional | Accessible in templates as `node.ColumnName` |

### Device Sheets

- Every non-Node device sheet is stored as a **list of dicts**, regardless of row count
- Single-row access: `{% set v = node.VPRN1000[0] %}` then `{{ v.Port_ID }}`
- Multi-row access: `{% for row in node.P2P %}{{ row.Port_ID }}{% endfor %}`
- Empty check: `{% if node.VPRN1000 %}` — empty list `[]` evaluates to `False`

### Common Parameters

Place shared values **above** the `NodeName` header row in any device sheet:

```
Row 2:  (empty)   VPRN_ID   SAP_IN   Customer_ID    ← common header
Row 3:  (empty)   1000      1000     1000            ← common values
Row 4:  (empty)
Row 5:  NodeName  Port_ID   InterfaceIP  FW_AS       ← device data starts here
Row 6:  R10       1/1/c12/1 100.30.1.18  65100
```

Common parameters are automatically merged into every device row. In the template:

```jinja2
{% set v = node.VPRN1000[0] %}
vprn {{ v.VPRN_ID }}        {# ← from common params #}
    port {{ v.Port_ID }}    {# ← from device data   #}
```

### Global Sheets

Sheets **without** a `NodeName` column are treated as global — shared across all devices.

```
Sheet: SERVER
NTP1        NTP2        Radius1     Syslog1
1.1.1.1     2.2.2.2     3.3.3.3     4.4.4.4
```

Template access:

```jinja2
server {{ SERVER.NTP1 }}
server {{ SERVER.NTP2 }}
server {{ SERVER.Radius1 }} secret <key>
```

---

## 📝 Jinja2 Template Reference

### Available Variables

| Variable | Type | Description |
|----------|------|-------------|
| `node` | `dict` | Current device data. Node sheet fields + all service sheets as lists |
| `node_list` | `list of dicts` | All devices. Use for BGP neighbor loops etc. |
| `node_lookup` | `dict of dicts` | All devices indexed by NodeName. Use for direct device lookup |
| `SERVER` / `QOS` / `POLICY` | `dict` | Global sheet data (one variable per global sheet) |

### Common Patterns

**Check if a device has a service:**
```jinja2
{% if node.VPRN1000 %}
    ... generate VPRN1000 config ...
{% endif %}
```

**Single-row sheet access:**
```jinja2
{% set v = node.VPRN1000[0] %}
vprn {{ v.VPRN_ID }} name "VPRN1000" customer {{ v.Customer_ID }} create
    interface "{{ v.InterfaceName }}" create
        address {{ v.InterfaceIP }}/31
        sap {{ v.Port_ID }}:{{ v.VLAN }} create
        exit
    exit
exit
```

**Multi-row sheet loop:**
```jinja2
{% for intf in node.P2P %}
interface "{{ intf.Inter_Name }}"
    address {{ intf.Inter_IP }}/31
    port {{ intf.Port_ID }}
exit
{% endfor %}
```

**Loop with inline filter:**
```jinja2
{% for row in node.VPRN1001 if row.Port_ID %}
    sap {{ row.Port_ID }}:{{ row.VLAN }} create
{% endfor %}
```

**Column (Key-Value) loop for multi-port single-row sheets:**
```jinja2
{% set v = node.VPLS2600[0] %}
{% for key, val in v.items() if key.startswith('Port') and val %}
sap {{ val }}:{{ v.VLAN }} create
exit
{% endfor %}
```

**MDA configuration (column loop):**
```jinja2
{% set mda = node.MDA[0] %}
card {{ mda.IOM }}
    card-type {{ mda.IOM_Type }}
    {% for key, val in mda.items() if key.startswith('MDA') and val %}
    mda {{ key | replace('MDA', '') }}
        mda-type {{ val }}
        no shutdown
    exit
    {% endfor %}
exit
```

**BGP Route Reflector vs Client:**
```jinja2
{% if node.IBGP_Role == 'RR' %}
    group "RR-to-Clients"
        {% for client in node_list if client.IBGP_Role != 'RR' %}
        neighbor {{ client.SystemIP }}
        exit
        {% endfor %}
    exit
{% else %}
    group "Client-to-RR"
        {% for rr in node_list if rr.IBGP_Role == 'RR' %}
        neighbor {{ rr.SystemIP }}
        exit
        {% endfor %}
    exit
{% endif %}
```

**Direct lookup of another device:**
```jinja2
neighbor {{ node_lookup.R2.SystemIP }}
neighbor {{ node_lookup['R1'].SystemIP }}
```

**P2P connector deduplication:**
```jinja2
{% set seen = namespace(connectors=[]) %}
{% for intf in node.P2P %}
{% if intf.Connector and intf.Connector not in seen.connectors %}
{% set seen.connectors = seen.connectors + [intf.Connector] %}
connector {{ intf.Connector }}
    breakout {{ intf.Break_Out }}
exit
{% endif %}
{% endfor %}
```

**Include sub-templates:**
```jinja2
{% include "7250-IXR-R6D/1_device_7250r6d_system.j2" %}
{% include "7250-IXR-R6D/8_device_7250r6d_Router_IGP.j2" %}
{% include "7250-IXR-R6D/9_device_7250r6d_iBGP.j2" %}
```

---

## ⚠️ Common Mistakes

| Mistake | Wrong | Correct |
|---------|-------|---------|
| Variable name case | `{{ Node.SystemIP }}` | `{{ node.SystemIP }}` |
| Space before mask | `{{ v.IP }} /31` | `{{ v.IP }}/31` |
| Variable in quotes | `vpls "row.Name"` | `vpls "{{ row.Name }}"` |
| Wrong sheet in loop | `if node.VPRN1001` / `for r in node.VPRN1000` | `if node.VPRN1001` / `for r in node.VPRN1001` |
| Include path | `{% include '../templates/igp.j2' %}` | `{% include '7250-IXR-R6D/igp.j2' %}` |

---

## 🔧 Troubleshooting

| Error | Likely Cause |
|-------|-------------|
| `'Node' is undefined` | Template uses capital `N` — replace `Node.` with `node.` |
| `TemplateNotFound` | `TemplateName` in Excel doesn't match actual `.j2` filename |
| `KeyError: 'FieldName'` | Column name in Excel doesn't match variable name in template |
| `object is not iterable` | Using `{% for %}` on wrong data — check sheet type |
| `neighbor` empty in BGP | Column name typo e.g. `FW_Inter_IP` vs `FW_Inster_IP` |
| Missing NodeName warning | Device in Node sheet has no `TemplateName` value |

---

## 🖥️ Supported Platforms

| Platform | Template |
|----------|----------|
| Nokia 7250 IXR-R6D | `7250-IXR-ALL.j2` |
| Nokia 7250 IXR-E2C | `7250-IXR-ALL.j2` |
| Nokia 7705 SAR-8 | `7250-IXR-ALL.j2` |

> Additional platforms can be supported by creating new Jinja2 templates and setting the `TemplateName` column in the Node sheet accordingly.

---

## 📦 Requirements

```
Python    >= 3.10
pandas    >= 2.0.0
openpyxl  >= 3.1.0
jinja2    >= 3.1.0
pydantic  >= 2.0.0
```

Install all dependencies:

```bash
pip install pandas openpyxl jinja2 pydantic
```

Or using `requirements.txt`:

```bash
pip install -r requirements.txt
```

---

## 📄 License

MIT License — free to use, modify, and distribute.

---
