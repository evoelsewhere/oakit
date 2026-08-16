const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const EMPTY_SHAPE_TREE =
  '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>';

const MINIMAL_SLIDE_MASTER =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<p:sldMaster xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIPS_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}">` +
  `<p:cSld name="OAKit Minimal Master">${EMPTY_SHAPE_TREE}</p:cSld>` +
  '<p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>' +
  '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
  '</p:sldMaster>';

const MINIMAL_SLIDE_LAYOUT =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<p:sldLayout xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIPS_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}" type="blank" preserve="1">` +
  `<p:cSld name="Blank">${EMPTY_SHAPE_TREE}</p:cSld>` +
  '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
  '</p:sldLayout>';

export function serializeMinimalSlideMaster(): string {
  return MINIMAL_SLIDE_MASTER;
}

export function serializeMinimalSlideLayout(): string {
  return MINIMAL_SLIDE_LAYOUT;
}
