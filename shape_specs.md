# Shape Specifications

Standard shape definitions for PPTX graphics.

## Basic Shapes

| Shape Type | MSO_SHAPE | Use Case |
|------------|-----------|----------|
| Rounded Rectangle | `MSO_SHAPE.ROUNDED_RECTANGLE` | Process nodes, general boxes |
| Rectangle | `MSO_SHAPE.RECTANGLE` | Headers, data containers |
| Diamond | `MSO_SHAPE.DIAMOND` | Decision points |
| Cylinder | `MSO_SHAPE.CAN` | Databases, storage |
| Ellipse | `MSO_SHAPE.OVAL` | Start/End points |
| Triangle | `MSO_SHAPE.ISOSCELES_TRIANGLE` | Direction indicators |

## Connector Types

| Type | MSO_CONNECTOR | Arrow |
|------|---------------|-------|
| Straight | `MSO_CONNECTOR.STRAIGHT` | End arrowhead |
| Elbow | `MSO_CONNECTOR.ELBOW` | End arrowhead |
| Curve | `MSO_CONNECTOR.CURVE` | End arrowhead |

## Standard Sizes

- Small box: 2.0" x 0.8"
- Medium box: 2.5" x 1.0"
- Large box: 3.0" x 1.5"
- Layer banner: 12.0" x 1.2"

## Font Guidelines

- Title: 32-44pt, bold
- Heading: 18-24pt, bold
- Body: 14-16pt
- Labels: 10-12pt
- Minimum: 8pt (never go below)
