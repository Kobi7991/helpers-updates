#target illustrator

// Ensure $.global exists before defining functions
(function() {
  try {
    if (typeof $.global === 'undefined') {
      $.global = {};
    }
  } catch (e) {
    // If we can't set $.global, try alternative
    try {
      if (typeof $ !== 'undefined' && typeof $.global === 'undefined') {
        $.global = {};
      }
    } catch (e2) {}
  }
})();

function helpersPing() {
    try {
        var doc = app.documents.length ? app.activeDocument : null;
        if (!doc) return "OK (no document open)";
        return "OK (doc: " + doc.name + ")";
    } catch (e) {
        return "ERROR: " + e;
    }
}
// Explicitly attach to $.global for CEP access
try { if (typeof $.global !== 'undefined') $.global.helpersPing = helpersPing; } catch(e) {}

function helpersAlert(msg) {
    try {
        alert(msg);
        return "OK";
    } catch (e) {
        return "ERROR: " + e;
    }
}

function helpersCleanup() {
    try {
        if (!app.documents.length) return "No document open.";
        var doc = app.activeDocument;
        var deletedLayers = [];
        
        // Remove Helpers layer (Logo Grid, Base Grid, etc.)
        var helpersLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === "Helpers") { 
                helpersLayer = doc.layers[i]; 
                break; 
            }
        }
        if (helpersLayer) {
            helpersLayer.remove();
            deletedLayers.push("Helpers");
        }
        
        // Remove Helpersub_Clearspace layer (Clear Space output)
        var clearspaceLayer = null;
        for (var j = 0; j < doc.layers.length; j++) {
            if (doc.layers[j].name === "Helpersub_Clearspace") { 
                clearspaceLayer = doc.layers[j]; 
                break; 
            }
        }
        if (clearspaceLayer) {
            clearspaceLayer.remove();
            deletedLayers.push("Helpersub_Clearspace");
        }
        
        if (deletedLayers.length === 0) {
            return "No layers found (nothing to delete).";
        }
        
        return "Deleted " + deletedLayers.join(" and ") + " layer(s).";
    } catch (e) {
        return "ERROR: " + e;
    }
}
// Explicitly attach to $.global for CEP access
try { if (typeof $.global !== 'undefined') $.global.helpersCleanup = helpersCleanup; } catch(e) {}

/**
 * Context-aware cleanup: remove only the elements for the given mode.
 * mode: "logo" = Logo Grid (Group_Anchors, Group_Handles, Group_Outline, Group_Gridlines)
 *       "clearspace" = Clearspace (Helpersub_Clearspace layer)
 *       "base" = Base Grid (Group_BaseGrid)
 * Does not delete other content on the page.
 */
function helpersCleanupByMode(mode) {
    try {
        if (!app.documents.length) return "No document open.";
        var doc = app.activeDocument;
        var modeStr = String(mode || "").toLowerCase();
        var removed = [];

        if (modeStr === "clearspace") {
            // Force delete all X-REF objects in the document first (any layer)
            var xrefMsg = helpersForceDeleteAllXRef();
            for (var j = 0; j < doc.layers.length; j++) {
                if (doc.layers[j].name === "Helpersub_Clearspace") {
                    doc.layers[j].remove();
                    removed.push("Helpersub_Clearspace");
                    break;
                }
            }
            if (removed.length === 0) {
                return xrefMsg.indexOf("Removed") === 0 ? xrefMsg : "No clearspace layer found (nothing to delete).";
            }
            return (xrefMsg.indexOf("Removed") === 0 ? xrefMsg + " " : "") + "Cleaned clearspace layer.";
        }

        if (modeStr === "logo" || modeStr === "base") {
            var helpersLayer = null;
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === "Helpers") {
                    helpersLayer = doc.layers[i];
                    break;
                }
            }
            if (!helpersLayer) {
                return "Layer not found (nothing to delete).";
            }

            var groupNames = modeStr === "logo"
                ? ["Group_Anchors", "Group_Handles", "Group_Outline", "Group_Gridlines"]
                : ["Group_BaseGrid"];

            for (var g = 0; g < groupNames.length; g++) {
                var grp = _findGroupInLayer(helpersLayer, groupNames[g]);
                if (grp) {
                    grp.remove();
                    removed.push(groupNames[g]);
                }
            }

            if (removed.length === 0) {
                return "No " + (modeStr === "logo" ? "logo grid" : "base grid") + " groups found (nothing to delete).";
            }

            // If Helpers layer has no pageItems left, remove the layer
            try {
                if (helpersLayer.pageItems.length === 0) {
                    helpersLayer.remove();
                }
            } catch (eL) {}

            return "Cleaned " + (modeStr === "logo" ? "logo grid" : "base grid") + ".";
        }

        return "Unknown mode. Use \"logo\", \"clearspace\", or \"base\".";
    } catch (e) {
        return "ERROR: " + e;
    }
}
try { if (typeof $.global !== 'undefined') $.global.helpersCleanupByMode = helpersCleanupByMode; } catch(e) {}

function helpersMakeGuides(target) {
    try {
        if (!app.documents.length) return "No document open.";
        var doc = app.activeDocument;

        // find Helpers layer
        var gridLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === "Helpers") { gridLayer = doc.layers[i]; break; }
        }
        if (!gridLayer) return "Layer not found.";

        gridLayer.locked = false;

        // Base Grid: stored as group Group_BaseGrid on Helpers layer (not a sublayer)
        var container = null;
        if (target === "Base Grid") {
            container = _findGroupInLayer(gridLayer, "Group_BaseGrid");
            if (!container) return "Base Grid group not found. Generate a Base Grid first.";
        } else {
            // find sublayer by name (e.g. Logo Grid guides stored on a sublayer)
            for (var j = 0; j < gridLayer.layers.length; j++) {
                if (gridLayer.layers[j].name === target) { container = gridLayer.layers[j]; break; }
            }
            if (!container) return "Sublayer '" + target + "' not found.";
        }

        if (container.locked !== undefined) container.locked = false;
        if (container.visible !== undefined) container.visible = true;

        var items = [];
        _collectItemsRecursive(container, items);

        var converted = 0;

        for (var k = 0; k < items.length; k++) {
            var it = items[k];
            try {
                if (it.locked) it.locked = false;
                if (it.hidden) it.hidden = false;

                if (it.typename === "PathItem") {
                    it.filled = false;
                    it.stroked = false;
                    it.guides = true;
                    converted++;
                } else if (it.typename === "CompoundPathItem") {
                    for (var p = 0; p < it.pathItems.length; p++) {
                        var pi = it.pathItems[p];
                        pi.filled = false;
                        pi.stroked = false;
                        pi.guides = true;
                        converted++;
                    }
                }
            } catch (inner) {}
        }

        return "Made guides from " + converted + " paths in '" + target + "'.";
    } catch (e) {
        return "ERROR: " + e;
    }
}
// Explicitly attach to $.global for CEP access
try { if (typeof $.global !== 'undefined') $.global.helpersMakeGuides = helpersMakeGuides; } catch(e) {}

// ---------- helpers ----------
function _rgb(r,g,b){
    var c = new RGBColor();
    c.red = r; c.green = g; c.blue = b;
    return c;
}


// ===== Added in v8.4.1a: missing helpers referenced by Base Grid / Clearspace =====
function _grayStroke(v){
  // returns an RGB gray stroke color (v can be 0..255). default = 160
  var g = (v===undefined || v===null) ? 160 : v;
  return _rgb(g,g,g);
}

function _addRect(layer, L, T, R, B, strokeW, strokeColor, opacity){
  // Rectangle from bounds [L,T,R,B]
  var w = R - L;
  var h = T - B;
  var r = layer.pathItems.rectangle(T, L, w, h);
  r.stroked = true;
  r.strokeWidth = (strokeW===undefined||strokeW===null) ? 1 : strokeW;
  r.strokeColor = strokeColor || _grayStroke();
  if(opacity!==undefined && opacity!==null) r.opacity = opacity;
  r.filled = false;
  return r;
}

// Collect PathItem objects from various selection types (non-destructive)
function _collectPathItems(item){
  var out = [];
  if (!item) return out;

  function walk(it){
    if (!it) return;
    var t = it.typename;
    if (t === 'PathItem'){
      try {
        // ignore guides and clipping masks
        if (it.guides) return;
        // clipping paths can still have useful anchors, but usually noise; skip by default
        if (it.clipping) return;
      } catch(e){}
      out.push(it);
      return;
    }
    if (t === 'CompoundPathItem'){
      try {
        for (var i=0; i<it.pathItems.length; i++){
          walk(it.pathItems[i]);
        }
      } catch(e){}
      return;
    }
    if (t === 'GroupItem' || t === 'Layer'){
      try {
        var items = it.pageItems;
        for (var j=0; j<items.length; j++){
          walk(items[j]);
        }
      } catch(e){}
      return;
    }
    // TextFrame is intentionally ignored to keep this non-destructive.
    // Users can outline text first if they need anchors/handles.
  }

  walk(item);
  return out;
}

// Add a filled triangular arrowhead at (x,y) pointing at angle (radians)
function _addArrowHead(group, x, y, angleRad, size, fillColor, opacity){
  try {
    var s = size || 6;
    var a = angleRad || 0;

    // Triangle in local coords pointing right: tip at (0,0), base at (-s, ±s*0.6)
    var tipX = x;
    var tipY = y;

    var bx = -s;
    var by = s * 0.6;

    function rot(px, py){
      var c = Math.cos(a), si = Math.sin(a);
      return [px*c - py*si, px*si + py*c];
    }

    var p1 = rot(0, 0);         // tip
    var p2 = rot(bx,  by);      // base top
    var p3 = rot(bx, -by);      // base bottom

    var item = group.pathItems.add();
    item.setEntirePath([
      [tipX + p1[0], tipY + p1[1]],
      [tipX + p2[0], tipY + p2[1]],
      [tipX + p3[0], tipY + p3[1]]
    ]);
    item.closed = true;
    item.stroked = false;
    item.filled = true;
    item.fillColor = fillColor;
    if (opacity !== undefined && opacity !== null) item.opacity = opacity;
    return item;
  } catch(e){
    return null;
  }
}
// ================================================================================

function _findLayer(doc, name) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) return doc.layers[i];
    }
    return null;
}

function _ensureLayer(doc, name) {
    var l = _findLayer(doc, name);
    if (!l) {
        l = doc.layers.add();
        l.name = name;
    }
    // Always keep our layer on top so generated grid elements appear above artwork
    try { l.visible = true; } catch (e1) {}
    try { l.locked = false; } catch (e2) {}
    try { l.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (e3) {}
    return l;
}

function _ensureSublayer(parentLayer, name) {
    var l = null;
    for (var i = 0; i < parentLayer.layers.length; i++) {
        if (parentLayer.layers[i].name === name) { l = parentLayer.layers[i]; break; }
    }
    if (!l) {
        l = parentLayer.layers.add();
        l.name = name;
    }
    try { l.visible = true; } catch (e1) {}
    try { l.locked = false; } catch (e2) {}
    return l;
}

function _clearLayerItems(layer) {
    // remove pageItems in reverse
    for (var i = layer.pageItems.length - 1; i >= 0; i--) {
        try { layer.pageItems[i].remove(); } catch(e){}
    }
    for (var j = layer.layers.length - 1; j >= 0; j--) {
        try { layer.layers[j].remove(); } catch(e){}
    }
}

// ===================== Live Preview Helpers =====================
// Find existing group by name within a layer or group
/**
 * Find a group by name within a container (Layer or GroupItem)
 * This function is critical for isolated updates - it only finds the exact group by name
 * Returns null if not found, ensuring no accidental modifications
 */
function _findGroupInLayer(container, groupName) {
    if (!container || !groupName) return null;
    try {
        // Check if container is a Layer
        if (container.typename === 'Layer') {
            // Search through all groupItems in the layer
            for (var i = 0; i < container.groupItems.length; i++) {
                try {
                    var group = container.groupItems[i];
                    // Exact name match - case sensitive for safety
                    if (group && group.name === groupName) {
                        return group;
                    }
                } catch(e) {
                    // Skip invalid items
                    continue;
                }
            }
        }
        // Check if container is a GroupItem
        else if (container.typename === 'GroupItem') {
            // Search through pageItems (which may include nested groups)
            for (var j = 0; j < container.pageItems.length; j++) {
                try {
                    var item = container.pageItems[j];
                    // Only match GroupItems with exact name
                    if (item && item.typename === 'GroupItem' && item.name === groupName) {
                        return item;
                    }
                } catch(e) {
                    // Skip invalid items
                    continue;
                }
            }
        }
    } catch(e) {
        // Return null on any error - fail safely
    }
    return null;
}

// Find existing sublayer by name within a parent layer
function _findSublayer(parentLayer, sublayerName) {
    if (!parentLayer) return null;
    try {
        for (var i = 0; i < parentLayer.layers.length; i++) {
            if (parentLayer.layers[i].name === sublayerName) {
                return parentLayer.layers[i];
            }
        }
    } catch(e) {}
    return null;
}

// Update stroke properties recursively on all path items in a group/layer
function _updateStrokeRecursive(item, strokeWidth, strokeColor, opacity) {
    if (!item) return;
    var tn = item.typename;
    
    if (tn === 'PathItem') {
        try {
            if (strokeWidth !== undefined && strokeWidth !== null) {
                item.strokeWidth = strokeWidth;
            }
            if (strokeColor) {
                item.strokeColor = strokeColor;
            }
            if (opacity !== undefined && opacity !== null) {
                item.opacity = opacity;
            }
        } catch(e) {}
        return;
    }
    
    if (tn === 'CompoundPathItem') {
        for (var i = 0; i < item.pathItems.length; i++) {
            _updateStrokeRecursive(item.pathItems[i], strokeWidth, strokeColor, opacity);
        }
        return;
    }
    
    if (tn === 'GroupItem') {
        for (var j = 0; j < item.pageItems.length; j++) {
            _updateStrokeRecursive(item.pageItems[j], strokeWidth, strokeColor, opacity);
        }
        return;
    }
}

// Update fill color and size for anchor markers (squares/circles)
function _updateAnchorMarkers(group, anchorSize, anchorShape, fillColor) {
    if (!group) return;
    try {
        for (var i = 0; i < group.pageItems.length; i++) {
            var item = group.pageItems[i];
            if (item.typename === 'PathItem') {
                // Check if it's a square or circle (anchors are filled)
                try {
                    if (item.filled && !item.stroked) {
                        // This is likely an anchor marker
                        var bounds = item.geometricBounds; // [top, left, bottom, right]
                        var width = bounds[2] - bounds[1];
                        var height = bounds[0] - bounds[3];
                        
                        // Update size if needed (recreate if size changed significantly)
                        var currentSize = Math.max(width, height);
                        if (Math.abs(currentSize - anchorSize) > 0.1) {
                            // Size changed, need to recreate - but we'll do this by clearing and redrawing
                            // For now, just update color
                            if (fillColor) {
                                item.fillColor = fillColor;
                            }
                        } else {
                            // Just update color
                            if (fillColor) {
                                item.fillColor = fillColor;
                            }
                        }
                    }
                } catch(e) {}
            } else if (item.typename === 'GroupItem') {
                _updateAnchorMarkers(item, anchorSize, anchorShape, fillColor);
            }
        }
    } catch(e) {}
}

/**
 * Clear all items from a specific group
 * CRITICAL: This function only clears the group passed to it - never touches other groups
 * Non-destructive: Only removes pageItems, preserves the group itself
 */
function _clearGroupItems(group) {
    if (!group) return;
    try {
        // Verify we have a valid group
        if (group.typename !== 'GroupItem') return;
        
        // Remove in reverse order to avoid index shifting issues
        var items = group.pageItems;
        for (var i = items.length - 1; i >= 0; i--) {
            try {
                var item = items[i];
                if (item) {
                    item.remove();
                }
            } catch(e) {
                // Continue removing other items even if one fails
                continue;
            }
        }
    } catch(e) {
        // Fail silently - don't break the update process
    }
}
// ===================== End Live Preview Helpers =====================

function _getBounds(doc, useSelection) {
    // returns [left, top, right, bottom]
    try {
        if (useSelection && doc.selection && doc.selection.length) {
            // union visibleBounds
            var b = doc.selection[0].visibleBounds; // [l,t,r,b]
            var l = b[0], t = b[1], r = b[2], bt = b[3];
            for (var i = 1; i < doc.selection.length; i++) {
                var bi = doc.selection[i].visibleBounds;
                l = Math.min(l, bi[0]);
                t = Math.max(t, bi[1]);
                r = Math.max(r, bi[2]);
                bt = Math.min(bt, bi[3]);
            }
            return [l,t,r,bt];
        }
    } catch(e) {}
    // active artboard
    var idx = doc.artboards.getActiveArtboardIndex();
    var rect = doc.artboards[idx].artboardRect; // [l,t,r,b]
    return [rect[0], rect[1], rect[2], rect[3]];
}

function _getActiveArtboardRect(doc) {
    var idx = doc.artboards.getActiveArtboardIndex();
    var rect = doc.artboards[idx].artboardRect; // [l,t,r,b]
    return [rect[0], rect[1], rect[2], rect[3]];
}

function _addLine(layer, x1,y1,x2,y2, strokeW, color) {
    var p = layer.pathItems.add();
    p.stroked = true;
    p.filled = false;
    p.strokeWidth = strokeW;
    p.strokeColor = color;
    p.setEntirePath([[x1,y1],[x2,y2]]);
    p.closed = false;
    try { p.selected = false; } catch(e) {}
    return p;
}

// Square marker centered at (cx, cy)
function _addSquare(layer, cx, cy, size, strokeW, strokeColor, fillColor, opacity) {
    var s = size;
    var top = cy + (s/2);
    var left = cx - (s/2);
    var r = layer.pathItems.rectangle(top, left, s, s);
    r.stroked = true;
    r.strokeWidth = strokeW;
    r.strokeColor = strokeColor;
    r.filled = true;
    r.fillColor = fillColor;
    try { if (opacity !== undefined && opacity !== null) r.opacity = opacity; } catch(e) {}
    try { r.selected = false; } catch(e) {}
    return r;
}

function _intersectionsWithRect(nx, ny, c, L, T, R, B) {
    // line: nx*x + ny*y = c
    var pts = [];

    function addPt(x,y){
        // dedupe approx
        for (var i=0;i<pts.length;i++){
            if (Math.abs(pts[i][0]-x) < 0.01 && Math.abs(pts[i][1]-y) < 0.01) return;
        }
        pts.push([x,y]);
    }

    // vertical edges x=L and x=R
    if (Math.abs(ny) > 1e-6) {
        var yL = (c - nx*L)/ny;
        if (yL <= T+0.01 && yL >= B-0.01) addPt(L,yL);
        var yR = (c - nx*R)/ny;
        if (yR <= T+0.01 && yR >= B-0.01) addPt(R,yR);
    }
    // horizontal edges y=T and y=B
    if (Math.abs(nx) > 1e-6) {
        var xT = (c - ny*T)/nx;
        if (xT >= L-0.01 && xT <= R+0.01) addPt(xT,T);
        var xB = (c - ny*B)/nx;
        if (xB >= L-0.01 && xB <= R+0.01) addPt(xB,B);
    }
    return pts;
}

function _addParallelFamily(layer, angleDeg, spacing, bounds, strokeW, color) {
    var L=bounds[0], T=bounds[1], R=bounds[2], B=bounds[3];
    var a = angleDeg * Math.PI / 180.0;
    // direction d=(cos,sin), normal n=(-sin,cos)
    var nx = -Math.sin(a);
    var ny =  Math.cos(a);

    var corners = [
        [L,T],[R,T],[R,B],[L,B]
    ];
    var minC =  1e100;
    var maxC = -1e100;
    for (var i=0;i<corners.length;i++){
        var c = nx*corners[i][0] + ny*corners[i][1];
        minC = Math.min(minC, c);
        maxC = Math.max(maxC, c);
    }
    var start = Math.floor(minC/spacing)*spacing;
    var end   = Math.ceil(maxC/spacing)*spacing;

    for (var v=start; v<=end+0.0001; v+=spacing){
        var pts = _intersectionsWithRect(nx,ny,v,L,T,R,B);
        if (pts.length >= 2) {
            _addLine(layer, pts[0][0], pts[0][1], pts[1][0], pts[1][1], strokeW, color);
        }
    }
}


function _addHexGrid(layer, radius, bounds, strokeW, color) {
    // True hex tiling (pointy-top) inside bounds (no gaps).
    // radius = distance from center to a vertex.
    var L=bounds[0], T=bounds[1], R=bounds[2], B=bounds[3];
    var r = Number(radius);
    if (!r || r <= 0) r = 40;

    // Pointy-top hex geometry
    var hexW = Math.sqrt(3) * r;   // width across flats
    var hexH = 2 * r;              // height point-to-point
    var stepX = hexW;              // center-to-center in X
    var stepY = 1.5 * r;           // center-to-center in Y
    var offsetX = hexW / 2;        // odd row offset

    // Start slightly outside bounds so edges are covered
    var startX = L - 2*hexW;
    var startY = T + 2*hexH;

    var row = 0;
    for (var cy = startY; cy >= B - 2*hexH; cy -= stepY) {
        var ox = (row % 2) ? offsetX : 0;
        for (var cx = startX + ox; cx <= R + 2*hexW; cx += stepX) {

            // Quick reject if completely outside
            if (cx + hexW/2 < L || cx - hexW/2 > R || cy + r < B || cy - r > T) continue;

            var pts = [];
            for (var k=0; k<6; k++) {
                // pointy-top, rotated for nicer look
                var ang = (Math.PI/180) * (60*k - 30);
                pts.push([cx + r*Math.cos(ang), cy + r*Math.sin(ang)]);
            }

            var p = layer.pathItems.add();
            p.stroked = true;
            p.filled = false;
            p.strokeWidth = strokeW;
            p.strokeColor = color;
            p.closed = true;

            p.setEntirePath([
                [pts[0][0], pts[0][1]],
                [pts[1][0], pts[1][1]],
                [pts[2][0], pts[2][1]],
                [pts[3][0], pts[3][1]],
                [pts[4][0], pts[4][1]],
                [pts[5][0], pts[5][1]]
            ]);
        }
        row++;
    }
}


// --- Golden Ratio / Fibonacci helpers (Base Grid) ---
function _addRectOutline(layer, left, top, width, height, strokeW, color, opacity) {
    var r = layer.pathItems.rectangle(top, left, width, height);
    r.stroked = true;
    r.filled = false;
    r.strokeWidth = strokeW;
    r.strokeColor = color;
    if (opacity !== undefined && opacity !== null) { try { r.opacity = opacity; } catch(e){} }
    return r;
}

// Quarter-circle arc using a cubic Bezier approximation (kappa).
// quadrant: 0=TR, 1=TL, 2=BL, 3=BR (clockwise)
function _addQuarterArc(layer, cx, cy, radius, quadrant, strokeW, color, opacity) {
    var k = 0.5522847498307936; // (4/3)*tan(pi/8)
    var r = radius;

    var p0, p1, p2, p3;
    if (quadrant === 0) { // TR: from top to right
        p0 = [cx, cy + r];
        p3 = [cx + r, cy];
        p1 = [cx + k*r, cy + r];
        p2 = [cx + r, cy + k*r];
    } else if (quadrant === 1) { // TL: from left to top
        p0 = [cx - r, cy];
        p3 = [cx, cy + r];
        p1 = [cx - r, cy + k*r];
        p2 = [cx - k*r, cy + r];
    } else if (quadrant === 2) { // BL: from bottom to left
        p0 = [cx, cy - r];
        p3 = [cx - r, cy];
        p1 = [cx - k*r, cy - r];
        p2 = [cx - r, cy - k*r];
    } else { // 3 BR: from right to bottom
        p0 = [cx + r, cy];
        p3 = [cx, cy - r];
        p1 = [cx + r, cy - k*r];
        p2 = [cx + k*r, cy - r];
    }

    var path = layer.pathItems.add();
    path.stroked = true;
    path.filled = false;
    path.strokeWidth = strokeW;
    path.strokeColor = color;
    if (opacity !== undefined && opacity !== null) { try { path.opacity = opacity; } catch(e){} }

    path.setEntirePath([p0, p3]);
    path.pathPoints[0].leftDirection = p0;
    path.pathPoints[0].rightDirection = p1;
    path.pathPoints[1].leftDirection = p2;
    path.pathPoints[1].rightDirection = p3;

    path.closed = false;
    return path;
}

function _fitGoldenRectToBounds(bounds) {
    var L=bounds[0], T=bounds[1], R=bounds[2], B=bounds[3];
    var W = R - L;
    var H = T - B;
    var phi = (1 + Math.sqrt(5)) / 2;

    var w, h;
    if (W / H >= phi) {
        h = H;
        w = H * phi;
    } else {
        w = W;
        h = W / phi;
    }

    var left = L + (W - w) / 2;
    var top  = T - (H - h) / 2;
    return { left:left, top:top, width:w, height:h, phi:phi };
}

function _buildFibonacciSquaresInGoldenRect(rect, maxSteps) {
    // Returns array of squares: {left, top, size, dir}
    // dir indicates where the square is placed/cut from: 0=LEFT,1=TOP,2=RIGHT,3=BOTTOM
    // This creates the classic golden-rectangle tiling used for the golden spiral.
    var squares = [];
    var left = rect.left;
    var top = rect.top;
    var w = rect.width;
    var h = rect.height;
    var dir = 0;

    for (var i = 0; i < maxSteps; i++) {
        var size = Math.min(w, h);
        if (size <= 0.5) break;

        var sq = { left: left, top: top, size: size, dir: dir };

        if (dir === 0) { // LEFT
            // square occupies the left side
            sq.left = left;
            sq.top = top;
            left += size;
            w -= size;
        } else if (dir === 1) { // TOP
            // square occupies the top side
            sq.left = left;
            sq.top = top;
            top -= size;
            h -= size;
        } else if (dir === 2) { // RIGHT
            // square occupies the right side
            sq.left = left + (w - size);
            sq.top = top;
            w -= size;
        } else { // BOTTOM
            // square occupies the bottom side
            sq.left = left;
            sq.top = top - (h - size);
            h -= size;
        }

        squares.push(sq);
        dir = (dir + 1) % 4;
    }
    return squares;
}
function helpersGenerateBaseGrid(mode, gridSize, strokeW, useSelection, updateMode) {
    try {
        if (!app.documents.length) return "No document open.";
        var doc = app.activeDocument;

        var size = Number(gridSize);
        if (!size || size <= 0) size = 40;
        var sw = Number(strokeW);
        if (!sw || sw <= 0) sw = 0.5;

        var bounds = _getBounds(doc, !!useSelection);
        var L=bounds[0], T=bounds[1], R=bounds[2], B=bounds[3];

        // Layer setup - create individual group directly in root layer
        var root = _ensureLayer(doc, "Helpers");
        
        var isUpdate = !!updateMode;
        var groupName = 'Group_BaseGrid';
        
        // Find existing group or create new one
        var gridGroup = isUpdate ? _findGroupInLayer(root, groupName) : null;
        if (gridGroup) {
            // Clear existing grid for update
            _clearGroupItems(gridGroup);
        } else {
            gridGroup = root.groupItems.add();
            gridGroup.name = groupName;
        }
        
        // Always bring to front for visibility
        try { gridGroup.zOrder(ZOrderMethod.BRINGTOFRONT); } catch(e){}

        // style
        var color = _rgb(120, 190, 255); // light blue
        // slightly transparent look via opacity on group? simplest: set on each path
        // Illustrator scripting: PathItem.opacity exists
        function setOpacityAll(layer, val){
            for (var i=0;i<layer.pathItems.length;i++){
                try { layer.pathItems[i].opacity = val; } catch(e){}
            }
        }

        if (mode === "square") {
            // vertical lines
            for (var x = Math.floor(L/size)*size; x <= R+0.0001; x += size) {
                _addLine(gridGroup, x, T, x, B, sw, color);
            }
            // horizontal lines
            for (var y = Math.ceil(T/size)*size; y >= B-0.0001; y -= size) {
                _addLine(gridGroup, L, y, R, y, sw, color);
            }
            setOpacityAll(gridGroup, 35);
            return "Square grid created.";
        }

        if (mode === "isometric") {
            // Artboard as bounding box: lines extend precisely to artboard edges (no gaps at corners)
            var abBounds = _getActiveArtboardRect(doc);
            _addParallelFamily(gridGroup, 90,   size, abBounds, sw, color);
            _addParallelFamily(gridGroup, 150,  size, abBounds, sw, color);
            _addParallelFamily(gridGroup, 30, size, abBounds, sw, color);
            setOpacityAll(gridGroup, 25);
            return "Isometric grid created.";
        }

        
        if (mode === "hex") {
            // True hex grid (tiling)
            _addHexGrid(gridGroup, size, bounds, sw, color);
            setOpacityAll(gridGroup, 18);
            return "Hex grid created.";
        }

if (mode === "golden") {
            // Golden Ratio (SVG-accurate) — Squares + Circles ONLY (no arcs)
            // Based on SVG viewBox: 0 0 117.23 72.59

            var VB_W = 117.23;
            var VB_H = 72.59;

            // Outer frame (inset like the SVG stroke sits): x=0.24 y=0.24 w=VB_W-0.48 h=VB_H-0.48
            var OUTER = { x: 0.24, y: 0.24, w: (VB_W - 0.48), h: (VB_H - 0.48) };

            // Exact squares from your SVG (x,y from top-left, y goes down)
            var SQUARES = [
                { x: 0.24,  y: 0.24,  s: 72.11 }, // big left square
                { x: 72.35, y: 0.24,  s: 44.64 }, // big right square
                { x: 89.52, y: 44.88, s: 27.47 },
                { x: 72.35, y: 55.18, s: 17.17 },
                { x: 72.35, y: 44.88, s: 10.30 },
                { x: 82.65, y: 44.88, s: 6.87  },
                { x: 82.65, y: 51.75, s: 3.43  },
                { x: 86.08, y: 51.75, s: 3.43  }
            ];

            // Fit the viewBox into bounds (keep aspect, center)
            var L = bounds[0], T = bounds[1], R = bounds[2], B = bounds[3];
            var targetW = R - L;
            var targetH = T - B;
            var scale = Math.min(targetW / VB_W, targetH / VB_H);
            var fittedW = VB_W * scale;
            var fittedH = VB_H * scale;
            var cx = (L + R) / 2;
            var cy = (T + B) / 2;
            var fitLeft = cx - fittedW / 2;
            var fitTop  = cy + fittedH / 2;

            function mapX(x) { return fitLeft + x * scale; }
            function mapTopY(y) { return fitTop - y * scale; }

            // FIX: Use existing group in update mode, don't create new one
            var g = null;
            if (isUpdate) {
                // Update mode: Find and clear existing group
                g = _findGroupInLayer(root, groupName);
                if (g) {
                    _clearGroupItems(g);
                } else {
                    // Group doesn't exist yet, create it
                    g = root.groupItems.add();
                    g.name = groupName;
                }
            } else {
                // Initial generation: Create new group
                g = root.groupItems.add();
                g.name = groupName;
            }
            
            // Bring to front for visibility
            try { g.zOrder(ZOrderMethod.BRINGTOFRONT); } catch(e){}

            // Styling (kept subtle like reference)
            var gridCol = _rgb(140, 150, 170);
            var opGrid = 35;

            if (true) {
                var outerRect = _addRectOutline(
                    root,
                    mapX(OUTER.x),
                    mapTopY(OUTER.y),
                    OUTER.w * scale,
                    OUTER.h * scale,
                    sw,
                    gridCol,
                    opGrid
                );
                outerRect.move(g, ElementPlacement.PLACEATEND);
            }

            // Squares + inscribed circles
            for (var i = 0; i < SQUARES.length; i++) {
                var q = SQUARES[i];
                var left = mapX(q.x);
                var top = mapTopY(q.y);
                var size = q.s * scale;

                var sq = _addRectOutline(root, left, top, size, size, sw, gridCol, opGrid);
                sq.move(g, ElementPlacement.PLACEATEND);

                var ccx = left + size / 2;
                var ccy = top - size / 2;
                var c = _addCircle(root, ccx, ccy, size / 2, sw, gridCol, null, false, opGrid);
                c.move(g, ElementPlacement.PLACEATEND);
            }

            return "Golden ratio grid created.";
        }

        return "Unknown mode: " + mode;
    } catch (e) {
        return "ERROR: " + e;
    }
}
// Explicitly attach to $.global for CEP access
try { if (typeof $.global !== 'undefined') $.global.helpersGenerateBaseGrid = helpersGenerateBaseGrid; } catch(e) {}

// ---------- Logo Grid Generator ----------
function _addCircle(layer, cx, cy, radius, strokeW, strokeColor, fillColor, filled, opacity) {
    var d = radius * 2;
    var top = cy + radius;
    var left = cx - radius;
    var c = layer.pathItems.ellipse(top, left, d, d);
    c.stroked = true;
    c.filled = !!filled;
    c.strokeWidth = strokeW;
    c.strokeColor = strokeColor;
    if (filled && fillColor) c.fillColor = fillColor;
    if (opacity !== undefined && opacity !== null) { try { c.opacity = opacity; } catch(e){} }
    try { c.selected = false; } catch(e) {}
    return c;
}

function _pushUnique(arr, v, eps) {
    for (var i=0;i<arr.length;i++) {
        if (Math.abs(arr[i]-v) <= eps) return;
    }
    arr.push(v);
}


function _dedupeSorted(arr, eps, descending) {
    if (!arr || arr.length <= 1) return arr;
    var out = [];
    for (var i=0; i<arr.length; i++) {
        var v = arr[i];
        if (!out.length) { out.push(v); continue; }
        var last = out[out.length-1];
        if (Math.abs(v - last) > eps) out.push(v);
    }
    // if descending was requested but input sorted descending, out already descending.
    return out;
}

function _collectItemsRecursive(item, out) {
    if (!item) return;
    var tn = item.typename;

    // If this is live text, convert the DUPLICATE to outlines so it won't cover the original artwork
    if (tn === 'TextFrame') {
        try {
            var outlined = item.createOutline();
            item.remove();
            _styleStrokeRecursive(outlined, strokeColor, strokeWidth, opacity);
        } catch (e) {
            // If outlining fails, fall back to styling whatever we can
        }
        return;
    }
    if (tn === 'GroupItem') {
        for (var i=0; i<item.pageItems.length; i++) {
            _collectItemsRecursive(item.pageItems[i], out);
        }
        return;
    }
    if (tn === 'CompoundPathItem') {
        for (var j=0; j<item.pathItems.length; j++) {
            out.push(item.pathItems[j]);
        }
        return;
    }
    if (tn === 'PathItem') {
        out.push(item);
        return;
    }
    // TextFrame / others ignored for anchors/handles (outlines can duplicate whole selection)
}

// Collect clean gridline coordinates from geometric bounds.
// This produces far more professional "construction lines" than using every anchor point.
function _collectBoundsRecursive(item, xs, ys, eps) {
    if (!item) return;
    var tn = item.typename;

    // Groups: recurse into children
    if (tn === 'GroupItem') {
        for (var i=0; i<item.pageItems.length; i++) {
            _collectBoundsRecursive(item.pageItems[i], xs, ys, eps);
        }
        return;
    }

    // Compound paths: collect each child path item bounds (outer + holes)
    if (tn === 'CompoundPathItem') {
        for (var j=0; j<item.pathItems.length; j++) {
            _collectBoundsRecursive(item.pathItems[j], xs, ys, eps);
        }
        return;
    }

    // Most drawable items expose geometricBounds
    try {
        if (item.geometricBounds && item.geometricBounds.length === 4) {
            var gb = item.geometricBounds; // [top, left, bottom, right]
            var top = gb[0], left = gb[1], bottom = gb[2], right = gb[3];
            _pushUnique(xs, left, eps);
            _pushUnique(xs, right, eps);
            _pushUnique(ys, top, eps);
            _pushUnique(ys, bottom, eps);
        }
    } catch(e) {}
}

// --- Weighted coord helpers (for cleaner Logo Grid gridlines) ---
function _qKey(v, eps) {
    // Quantize to eps grid so near-identical values collapse.
    // Avoid floating point drift by rounding to 1/eps steps.
    if (!eps || eps <= 0) eps = 0.5;
    var step = 1 / eps;
    return String(Math.round(v * step) / step);
}

function _addWeightedCoord(map, v, eps, w) {
    if (!map) return;
    var k = _qKey(v, eps);
    map[k] = (map[k] || 0) + (w || 1);
}

// Back-compat alias (older builds used _addW)
function _addW(map, v, eps, w) {
    _addWeightedCoord(map, v, eps, w);
}

function _collectBoundsWeightedRecursive(item, xw, yw, eps, w) {
    if (!item) return;
    var tn = item.typename;

    if (tn === 'GroupItem') {
        for (var i=0; i<item.pageItems.length; i++) {
            _collectBoundsWeightedRecursive(item.pageItems[i], xw, yw, eps, w);
        }
        return;
    }

    if (tn === 'CompoundPathItem') {
        for (var j=0; j<item.pathItems.length; j++) {
            _collectBoundsWeightedRecursive(item.pathItems[j], xw, yw, eps, w);
        }
        return;
    }

    try {
        if (item.geometricBounds && item.geometricBounds.length === 4) {
            var gb = item.geometricBounds; // [top, left, bottom, right]
            var top = gb[0], left = gb[1], bottom = gb[2], right = gb[3];
            _addWeightedCoord(xw, left, eps, w);
            _addWeightedCoord(xw, right, eps, w);
            _addWeightedCoord(yw, top, eps, w);
            _addWeightedCoord(yw, bottom, eps, w);
        }
    } catch(e) {}
}

function _pickTopWeighted(map, maxCount) {
    // Returns numeric values sorted ascending, from a weight-map.
    var arr = [];
    for (var k in map) {
        if (!map.hasOwnProperty(k)) continue;
        arr.push({ v: parseFloat(k), w: map[k] });
    }
    // weight desc, then value asc (stable looking)
    arr.sort(function(a,b){
        if (b.w !== a.w) return b.w - a.w;
        return a.v - b.v;
    });
    if (maxCount && arr.length > maxCount) arr.length = maxCount;
    // sort by value for drawing
    arr.sort(function(a,b){ return a.v - b.v; });
    var out = [];
    for (var i=0; i<arr.length; i++) out.push(arr[i].v);
    return out;
}

// Back-compat helper used by some builds.
// Returns a sorted list of coordinates from a weight-map, while also forcing
// inclusion of key coordinates (e.g., selection bounds + centerline).
function _pickWeightedCoords(map, tol, maxCount, includeAlways) {
    if (!tol || tol <= 0) tol = 0.5;
    if (!maxCount || maxCount <= 0) maxCount = 12;

    var base = _pickTopWeighted(map, maxCount);
    var out = [];

    function pushUnique(v) {
        if (v === null || v === undefined) return;
        v = Number(v);
        if (isNaN(v)) return;
        for (var i = 0; i < out.length; i++) {
            if (Math.abs(out[i] - v) <= tol) return;
        }
        out.push(v);
    }

    for (var i = 0; i < base.length; i++) pushUnique(base[i]);
    if (includeAlways && includeAlways.length) {
        for (var j = 0; j < includeAlways.length; j++) pushUnique(includeAlways[j]);
    }

    out.sort(function(a, b) { return a - b; });
    return out;
}

function _styleStrokeRecursive(item, sw, strokeColor, opacity) {
    if (!item) return;
    var tn = item.typename;

    // If this is live text, convert the DUPLICATE to outlines so it won't cover the original artwork
    if (tn === 'TextFrame') {
        try {
            var outlined = item.createOutline();
            item.remove();
            _styleStrokeRecursive(outlined, sw, strokeColor, opacity);
        } catch (e) {
            // If outlining fails, fall back to styling it as-is (rare)
            try {
                item.filled = false;
                item.stroked = true;
                item.strokeWidth = sw;
                item.strokeColor = strokeColor;
                item.opacity = opacity;
            } catch (e2) {}
        }
        return;
    }
    if (tn === 'GroupItem') {
        for (var i=0; i<item.pageItems.length; i++) {
            _styleStrokeRecursive(item.pageItems[i], sw, strokeColor, opacity);
        }
        return;
    }
    if (tn === 'CompoundPathItem') {
        for (var j=0; j<item.pathItems.length; j++) {
            _styleStrokeRecursive(item.pathItems[j], sw, strokeColor, opacity);
        }
        return;
    }
    if (tn === 'PathItem') {
        try {
            item.filled = false;
            item.stroked = true;
            item.strokeWidth = sw;
            item.strokeColor = strokeColor;
            if (opacity !== undefined && opacity !== null) item.opacity = opacity;
        } catch(e){}
        return;
    }
}

function _duplicateSelectionToContainer(doc, targetContainer) {
    var sel = doc.selection;
    var dupes = [];
    if (!sel || sel.length === 0) return dupes;
    for (var i=0; i<sel.length; i++) {
        try {
            var d = sel[i].duplicate(targetContainer, ElementPlacement.PLACEATEND);
            dupes.push(d);
        } catch(e){}
    }
    return dupes;
}

function _timestampTag(){
    function pad(n){ return (n < 10 ? '0' : '') + n; }
    var d = new Date();
    return d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate()) + '_' + pad(d.getHours()) + '-' + pad(d.getMinutes()) + '-' + pad(d.getSeconds());
}

function _createRunLayer(doc, prefix){
    var l = doc.layers.add();
    l.name = prefix + ' ' + _timestampTag();
    try { l.zOrder(ZOrderMethod.BRINGTOFRONT); } catch(e){}
    return l;
}

// anchorShape is optional ("square" | "circle"); default is "square".
// handleShape is optional ("square" | "circle"); default is "circle".
// anchorSize and handleSize are COMPLETELY INDEPENDENT - changing one never affects the other.
// Strokes are per-feature (anchors/handles/outline/gridlines) so each module can have its own weight.
// If updateMode is true, will find and update existing groups instead of creating new ones.
// Each action creates its own top-level group for easy toggling: Group_Anchors, Group_Handles, Group_Outline, Group_Gridlines
// CRITICAL: Each feature operates in complete isolation - updating one feature never affects others
function helpersGenerateLogoGrid(
    doAnchors, doHandles, doOutlines, doGridlines,
    strokeAnchors, strokeHandles, strokeOutlines, strokeGridlines,
    anchorSize, handleSize, gridDir, anchorShape, handleShape, anchorFill, handleFill, gridType, updateMode,
    colorStrokeR, colorStrokeG, colorStrokeB,
    colorAnchorsR, colorAnchorsG, colorAnchorsB,
    colorHandlesR, colorHandlesG, colorHandlesB
){
    var doc = app.activeDocument;
    if (!doc) return 'No document';
    if (!doc.selection || doc.selection.length === 0) return 'Select at least 1 object';

    // Ensure Helpers layer exists
    var root = _ensureLayer(doc, 'Helpers');
    
    var isUpdate = !!updateMode;
    
    // Individual groups for each action - completely isolated
    // Each feature only touches its own group, never others
    var gAnch = null, gHand = null, gOut = null, gGrid = null;
    
    // ANCHORS: Only process if enabled, only touch Group_Anchors
    // CRITICAL ISOLATION: This block only affects Group_Anchors, never other groups
    if (doAnchors) {
        gAnch = _findGroupInLayer(root, 'Group_Anchors');
        if (gAnch && isUpdate) {
            // Update mode: Clear only this specific group's contents
            // This operation is completely isolated - only affects Group_Anchors
            _clearGroupItems(gAnch);
        } else if (!gAnch) {
            // Initial generation: Create new group only if it doesn't exist
            gAnch = root.groupItems.add();
            gAnch.name = 'Group_Anchors';
        }
        // Bring to front for visibility (only affects this group, doesn't modify others)
        // zOrder is safe - it only changes stacking order, not group membership
        if (gAnch) {
            try { gAnch.zOrder(ZOrderMethod.BRINGTOFRONT); } catch(e){}
        }
    }
    // If doAnchors is false, we don't touch Group_Anchors at all - it remains untouched
    
    // HANDLES: Only process if enabled, only touch Group_Handles
    // CRITICAL ISOLATION: This block only affects Group_Handles, never other groups
    if (doHandles) {
        gHand = _findGroupInLayer(root, 'Group_Handles');
        if (gHand && isUpdate) {
            _clearGroupItems(gHand);
        } else if (!gHand) {
            gHand = root.groupItems.add();
            gHand.name = 'Group_Handles';
        }
        if (gHand) {
            try { gHand.zOrder(ZOrderMethod.BRINGTOFRONT); } catch(e){}
    }
    }
    // If doHandles is false, we don't touch Group_Handles at all - it remains untouched
    
    // OUTLINES: Only process if enabled, only touch Group_Outline
    // CRITICAL ISOLATION: This block only affects Group_Outline, never other groups
    if (doOutlines) {
        gOut = _findGroupInLayer(root, 'Group_Outline');
        if (gOut && isUpdate) {
            _clearGroupItems(gOut);
        } else if (!gOut) {
            gOut = root.groupItems.add();
            gOut.name = 'Group_Outline';
        }
        if (gOut) {
            try { gOut.zOrder(ZOrderMethod.BRINGTOFRONT); } catch(e){}
    }
    }
    // If doOutlines is false, we don't touch Group_Outline at all - it remains untouched
    
    // GRIDLINES: Only process if enabled, only touch Group_Gridlines
    // CRITICAL ISOLATION: This block only affects Group_Gridlines, never other groups
    if (doGridlines) {
        gGrid = _findGroupInLayer(root, 'Group_Gridlines');
        if (gGrid && isUpdate) {
            _clearGroupItems(gGrid);
        } else if (!gGrid) {
            gGrid = root.groupItems.add();
            gGrid.name = 'Group_Gridlines';
        }
        if (gGrid) {
            try { gGrid.zOrder(ZOrderMethod.BRINGTOFRONT); } catch(e){}
        }
    }
    // If doGridlines is false, we don't touch Group_Gridlines at all - it remains untouched

    // Styling - use global colors if provided, otherwise fallback to gray
    var sc = (colorStrokeR !== undefined && colorStrokeG !== undefined && colorStrokeB !== undefined) 
        ? _rgb(Number(colorStrokeR) || 47, Number(colorStrokeG) || 140, Number(colorStrokeB) || 255)
        : _grayStroke();
    var scAnchors = (colorAnchorsR !== undefined && colorAnchorsG !== undefined && colorAnchorsB !== undefined)
        ? _rgb(Number(colorAnchorsR) || 47, Number(colorAnchorsG) || 140, Number(colorAnchorsB) || 255)
        : sc;
    var scHandles = (colorHandlesR !== undefined && colorHandlesG !== undefined && colorHandlesB !== undefined)
        ? _rgb(Number(colorHandlesR) || 47, Number(colorHandlesG) || 140, Number(colorHandlesB) || 255)
        : sc;
    var swAnch = Number(strokeAnchors) || 0.5;
    var swHand = Number(strokeHandles) || 0.5;
    var swOut  = Number(strokeOutlines) || 0.5;
    var swGrid = Number(strokeGridlines) || 0.5;

    // Outline (non-destructive): create stroke-only outlines WITHOUT leaving filled duplicates behind
    if (gOut) {
        var selForOutline = doc.selection;
        for (var oi = 0; oi < selForOutline.length; oi++) {
            try {
                var outlinedPI = _outlineItem(selForOutline[oi], sc, swOut);
                if (outlinedPI) {
                    try { outlinedPI.name = 'Outline'; } catch (eName) {}
                    try { outlinedPI.move(gOut, ElementPlacement.PLACEATEND); } catch (eMove) {}
                }
            } catch (eOut) {}
        }
    }

    // Gridlines (based on selection bounds)
    // Support multiple grid types (comma-separated)
    if (gGrid){
        try {
            var b = _getSelectionBounds(doc);
            // CRITICAL: Ensure gridType is a valid string, default to straightLinesGrid if missing
            var gridTypeStr = String(gridType || 'straightLinesGrid');
            // Split and process grid types - use compatible approach for ExtendScript
            var gridTypesArray = gridTypeStr.split(',');
            var gridTypes = [];
            for (var i = 0; i < gridTypesArray.length; i++) {
                var trimmed = String(gridTypesArray[i]).replace(/^\s+|\s+$/g, ''); // trim using regex (more compatible)
                if (trimmed && trimmed.length > 0) {
                    gridTypes.push(trimmed);
                }
            }
            
            // If no valid grid types after filtering, use default
            if (gridTypes.length === 0) {
                gridTypes = ['straightLinesGrid'];
            }
            
            // Draw each selected grid type
            for (var gt = 0; gt < gridTypes.length; gt++) {
                var currentGridType = gridTypes[gt];
                // Ensure currentGridType is valid before calling
                if (currentGridType && currentGridType.length > 0) {
                    _drawGridlinesMode(gGrid, b, sc, swGrid, gridDir, currentGridType);
                }
            }
        } catch(e) {
            // Error in gridlines execution - return error message
            return 'ERROR in Gridlines: ' + e.toString();
        }
    }

    // Anchors / Handles from path points
    var sel = doc.selection;
    for (var i=0; i<sel.length; i++){
        var item = sel[i];
        var paths = _collectPathItems(item);
        for (var p=0; p<paths.length; p++){
            var path = paths[p];
            var pts = path.pathPoints;
            for (var j=0; j<pts.length; j++){
                var pt = pts[j];
                if (gAnch){
                    _drawAnchor(gAnch, pt.anchor, anchorSize, anchorShape, scAnchors, swAnch, !!anchorFill);
                }
                if (gHand){
                    _drawHandle(gHand, pt, handleSize, handleShape, scHandles, swHand, !!handleFill);
                }
            }
        }
    }

    return 'OK';
}
// Explicitly attach to $.global for CEP access
try { if (typeof $.global !== 'undefined') $.global.helpersGenerateLogoGrid = helpersGenerateLogoGrid; } catch(e) {}

// ============================================
// Live Updates (Observer): in-place style/size only. No new objects, no full generation.
// CRITICAL SEPARATION: These functions are ONLY for live updates (slider/input changes).
// They NEVER call clean(), generate(), or remove() on groups. They only modify properties.
// Rule: Update ONLY the specific property (fillColor, strokeWidth, etc.) of found objects by name.
// If no objects are found (layer/group missing), do nothing. Do not delete or re-create.
// Generate button is the ONLY creator. Each updater only touches its tagged group.
// ============================================

/**
 * Dedicated shape swap function: Replace objects by name with new shape at same center.
 * NEVER calls clean() or generate(). Only removes and recreates individual items.
 * @param {GroupItem} group - The group containing the items to swap
 * @param {string} targetName - Name to match ("Anchor" or "HandleDot")
 * @param {string} newShape - "square" or "circle"
 * @param {number} size - Size for the new shape
 * @param {RGBColor} fillColor - Fill color (for anchors) or null for stroke-only
 * @param {RGBColor} strokeColor - Stroke color (for handles) or null for fill-only
 * @param {number} strokeWidth - Stroke width (for handles) or 0 for fill-only
 */
function _swapShape(group, targetName, newShape, size, fillColor, strokeColor, strokeWidth) {
    if (!group) return;
    var isSquare = (newShape + '').toLowerCase() === 'square';
    var s = Math.max(0.1, size);
    var items = [];
    // Copy to array to avoid live collection issues during iteration
    for (var ii = 0; ii < group.pageItems.length; ii++) {
        try {
            items.push(group.pageItems[ii]);
        } catch (e) {
            // Skip invalid items
        }
    }
    var i, item, gb, nm, cx, cy, left, top, newItem, storedOpacity;
    for (i = 0; i < items.length; i++) {
        item = items[i];
        if (!item) continue;
        try { 
            nm = item.name || ''; 
        } catch (e) { 
            nm = ''; 
        }
        if (nm !== targetName) continue;
        
        // Get geometric bounds to calculate center - MUST succeed
        try {
            gb = item.geometricBounds;
            if (!gb || gb.length < 4) continue;
        } catch (e) {
            continue;
        }
        
        // geometricBounds: [top, left, bottom, right] where top > bottom (Y increases upward)
        // Calculate center point: cx = (left + right) / 2, cy = (top + bottom) / 2
        cx = (gb[1] + gb[3]) / 2;  // X center: (left + right) / 2
        cy = (gb[0] + gb[2]) / 2;  // Y center: (top + bottom) / 2
        
        // Store opacity before removal
        try { 
            storedOpacity = item.opacity || 100; 
        } catch (e) { 
            storedOpacity = 100; 
        }
        
        // Convert center to top-left coordinates for rectangle/ellipse
        // Pattern matches _drawAnchor and _drawHandle: left = cx - s/2, top = cy + s/2
        left = cx - s / 2;
        top = cy + s / 2;
        
        // Remove the old item
        try {
            item.remove();
        } catch (e) {
            // If removal fails, skip this item (don't try to create new one)
            continue;
        }
        
        // Create new shape at exact same center point - MUST succeed or we've lost the item
        try {
            if (isSquare) {
                newItem = group.pathItems.rectangle(top, left, s, s);
            } else {
                newItem = group.pathItems.ellipse(top, left, s, s);
            }
            
            // Apply styles based on what was passed
            if (fillColor) {
                // Anchor: filled, no stroke
                newItem.filled = true;
                newItem.fillColor = fillColor;
                newItem.stroked = false;
            } else if (strokeColor) {
                // HandleDot: stroked, no fill
                newItem.filled = false;
                newItem.stroked = true;
                newItem.strokeColor = strokeColor;
                newItem.strokeWidth = strokeWidth || 0.5;
            }
            
            // Restore name and opacity
            try { newItem.name = targetName; } catch (e) {}
            try { newItem.opacity = storedOpacity; } catch (e) {}
        } catch (e) {
            // If creation fails, we've already removed the old item - this is a critical error
            // But we can't recover here, so just continue to next item
            // In production, you might want to log this error
        }
    }
}

function _scalePathAboutCenter(item, scaleFactor) {
    var pct = Math.max(1, Math.min(1000, scaleFactor * 100));
    try {
        item.resize(pct, pct, true, true, true, true, pct, Transformation.CENTER);
    } catch (e) {
        try { item.resize(pct, pct); } catch (e2) {}
    }
}

/**
 * Live Update: Replace Anchor shapes by name. NEVER calls clean() or generate().
 * Uses dedicated _swapShape() function for reliable Circle ↔ Square replacement.
 * Finds all items named "Anchor", swaps shape at exact same center with current size/color.
 */
function helpersUpdateLogoGridAnchors(colorR, colorG, colorB, anchorSize, anchorShape, anchorFill, anchorStrokeW) {
    try {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        var root = _findLayer(doc, 'Helpers');
        if (!root) return;
        var gAnch = _findGroupInLayer(root, 'Group_Anchors');
        if (!gAnch) return;
        var sc = _rgb(Number(colorR) || 47, Number(colorG) || 140, Number(colorB) || 255);
        var newSize = Math.max(0.1, Number(anchorSize) || 3);
        var shapeStr = String(anchorShape || 'square').replace(/['"]/g, '').trim().toLowerCase();
        var shape = (shapeStr === 'square') ? 'square' : 'circle';
        var useFill = (String(anchorFill) === 'true');
        var sw = Math.max(0.1, Number(anchorStrokeW) || 0.5);
        if (useFill) {
            _swapShape(gAnch, 'Anchor', shape, newSize, sc, null, 0);
        } else {
            _swapShape(gAnch, 'Anchor', shape, newSize, null, sc, sw);
        }
    } catch (e) {}
}
try { if (typeof $.global !== 'undefined') $.global.helpersUpdateLogoGridAnchors = helpersUpdateLogoGridAnchors; } catch(e) {}

/**
 * Live Update: Replace HandleDot shapes by name, update HandleLine stroke. NEVER calls clean() or generate().
 * HandleLine: Only updates strokeColor and strokeWidth (never removed).
 * HandleDot: Uses dedicated _swapShape() for reliable Circle ↔ Square replacement at same center.
 * Ensures handle dots stay connected to their lines during shape toggle.
 */
function helpersUpdateLogoGridHandles(colorR, colorG, colorB, handleSize, handleShape, strokeW, handleFill) {
    try {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        var root = _findLayer(doc, 'Helpers');
        if (!root) return;
        var gHand = _findGroupInLayer(root, 'Group_Handles');
        if (!gHand) return;
        var sc = _rgb(Number(colorR) || 47, Number(colorG) || 140, Number(colorB) || 255);
        var newSize = Math.max(0.1, Number(handleSize) || 3);
        var sw = Number(strokeW) || 0.5;
        var shapeStr = String(handleShape || 'circle').replace(/['"]/g, '').trim().toLowerCase();
        var shape = (shapeStr === 'square') ? 'square' : 'circle';
        var useFill = (String(handleFill) === 'true');
        var items = [];
        for (var ii = 0; ii < gHand.pageItems.length; ii++) items.push(gHand.pageItems[ii]);
        var i, item, nm;
        for (i = 0; i < items.length; i++) {
            item = items[i];
            if (!item) continue;
            try { nm = item.name || ''; } catch (e) { nm = ''; }
            if (nm === 'HandleLine') {
                try {
                    item.strokeColor = sc;
                    item.strokeWidth = sw;
                } catch (e) {}
            }
        }
        if (useFill) {
            _swapShape(gHand, 'HandleDot', shape, newSize, sc, null, 0);
        } else {
            _swapShape(gHand, 'HandleDot', shape, newSize, null, sc, sw);
        }
    } catch (e) {}
}
try { if (typeof $.global !== 'undefined') $.global.helpersUpdateLogoGridHandles = helpersUpdateLogoGridHandles; } catch(e) {}

function helpersUpdateLogoGridOutline(colorR, colorG, colorB, strokeW) {
    try {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        var root = _findLayer(doc, 'Helpers');
        if (!root) return;
        var gOut = _findGroupInLayer(root, 'Group_Outline');
        if (!gOut) return;
        var sc = _rgb(Number(colorR) || 47, Number(colorG) || 140, Number(colorB) || 255);
        var sw = Number(strokeW) || 0.5;
        var items = [];
        for (var ii = 0; ii < gOut.pageItems.length; ii++) items.push(gOut.pageItems[ii]);
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it) continue;
            try { if ((it.name || '') !== 'Outline') continue; } catch (e) { continue; }
            _updateStrokeRecursive(it, sw, sc, undefined);
        }
    } catch (e) {}
}
try { if (typeof $.global !== 'undefined') $.global.helpersUpdateLogoGridOutline = helpersUpdateLogoGridOutline; } catch(e) {}

function helpersUpdateLogoGridGridlines(colorR, colorG, colorB, strokeW) {
    try {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        var root = _findLayer(doc, 'Helpers');
        if (!root) return;
        var gGrid = _findGroupInLayer(root, 'Group_Gridlines');
        if (!gGrid) return;
        var sc = _rgb(Number(colorR) || 47, Number(colorG) || 140, Number(colorB) || 255);
        var sw = Number(strokeW) || 0.5;
        var items = [];
        for (var ii = 0; ii < gGrid.pathItems.length; ii++) items.push(gGrid.pathItems[ii]);
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            try {
                var nm = item.name || '';
                if (nm !== 'Gridline' && nm.indexOf('GRID_') !== 0) continue;
            } catch (e) { continue; }
            try {
                item.strokeColor = sc;
                item.strokeWidth = sw;
            } catch (e) {}
        }
    } catch (e) {}
}
try { if (typeof $.global !== 'undefined') $.global.helpersUpdateLogoGridGridlines = helpersUpdateLogoGridGridlines; } catch(e) {}

/**
 * Update Base Grid visuals only (stroke width and optional color). No deletion, no regeneration.
 * Finds Group_BaseGrid by name and sets stroke on existing path items only.
 */
function helpersUpdateBaseGridVisuals(strokeW, colorR, colorG, colorB) {
    try {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        var root = _findLayer(doc, 'Helpers');
        if (!root) return;
        var g = _findGroupInLayer(root, 'Group_BaseGrid');
        if (!g) return;
        var sw = Number(strokeW) || 0.5;
        var sc = (colorR !== undefined && colorG !== undefined && colorB !== undefined)
            ? _rgb(Number(colorR) || 120, Number(colorG) || 190, Number(colorB) || 255)
            : null;
        var items = [];
        for (var ii = 0; ii < g.pathItems.length; ii++) items.push(g.pathItems[ii]);
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            try {
                item.strokeWidth = sw;
                if (sc) item.strokeColor = sc;
            } catch (e) {}
        }
    } catch (e) {}
}
try { if (typeof $.global !== 'undefined') $.global.helpersUpdateBaseGridVisuals = helpersUpdateBaseGridVisuals; } catch(e) {}

// ============================================
// Akrivi Clearspace Generator (Brand Guideline System)
// ============================================

/**
 * Get geometric bounds of an object recursively (handles Groups, CompoundPaths, Text)
 * Returns {L, T, R, B} or null if invalid
 */
/**
 * Get bounds for a single item using visibleBounds (preferred) or geometricBounds (fallback)
 * visibleBounds includes stroke appearance and transformations
 * Returns: {left, top, right, bottom, width, height, centerX, centerY} or null
 */
function getBounds(item) {
    if (!item) return null;
    try {
        // Try visibleBounds first (includes stroke and transformations)
        var vb = null;
        try {
            vb = item.visibleBounds; // [left, top, right, bottom]
            if (vb && vb.length === 4 && !isNaN(vb[0]) && !isNaN(vb[1]) && !isNaN(vb[2]) && !isNaN(vb[3])) {
                var left = vb[0];
                var top = vb[1];
                var right = vb[2];
                var bottom = vb[3];
                var width = right - left;
                var height = top - bottom;
                if (width > 0 && height > 0) {
                    return {
                        left: left,
                        top: top,
                        right: right,
                        bottom: bottom,
                        width: width,
                        height: height,
                        centerX: left + width / 2,
                        centerY: bottom + height / 2
                    };
                }
            }
        } catch(e) {}
        
        // Fallback to geometricBounds
        try {
            var gb = item.geometricBounds; // [top, left, bottom, right]
            if (gb && gb.length === 4 && !isNaN(gb[0]) && !isNaN(gb[1]) && !isNaN(gb[2]) && !isNaN(gb[3])) {
                var left = gb[1];
                var top = gb[0];
                var right = gb[2];
                var bottom = gb[3];
                var width = right - left;
                var height = top - bottom;
                if (width > 0 && height > 0) {
                    return {
                        left: left,
                        top: top,
                        right: right,
                        bottom: bottom,
                        width: width,
                        height: height,
                        centerX: left + width / 2,
                        centerY: bottom + height / 2
                    };
                }
            }
        } catch(e) {}
    } catch(e) {}
    return null;
}

/**
 * Get combined bounds for multiple items (recursively handles groups)
 * Returns: {left, top, right, bottom, width, height, centerX, centerY} or null
 */
function getCombinedBounds(items) {
    if (!items || items.length === 0) return null;
    
    var minLeft = null;
    var maxTop = null;
    var maxRight = null;
    var minBottom = null;
    
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) continue;
        
        var bounds = null;
        try {
            var tn = item.typename;
            
            // For groups, recursively get bounds of all children
            if (tn === 'GroupItem') {
                var childItems = [];
                try {
                    for (var j = 0; j < item.pageItems.length; j++) {
                        childItems.push(item.pageItems[j]);
                    }
                } catch(e) {}
                bounds = getCombinedBounds(childItems);
            } else {
                // For single items, use getBounds
                bounds = getBounds(item);
            }
        } catch(e) {
            continue;
        }
        
        if (bounds) {
            if (minLeft === null || bounds.left < minLeft) minLeft = bounds.left;
            if (maxTop === null || bounds.top > maxTop) maxTop = bounds.top;
            if (maxRight === null || bounds.right > maxRight) maxRight = bounds.right;
            if (minBottom === null || bounds.bottom < minBottom) minBottom = bounds.bottom;
        }
    }
    
    if (minLeft !== null && maxTop !== null && maxRight !== null && minBottom !== null) {
        var width = maxRight - minLeft;
        var height = maxTop - minBottom;
        return {
            left: minLeft,
            top: maxTop,
            right: maxRight,
            bottom: minBottom,
            width: width,
            height: height,
            centerX: minLeft + width / 2,
            centerY: minBottom + height / 2
        };
    }
    
    return null;
}

/**
 * Legacy function for backward compatibility
 * @deprecated Use getBounds() or getCombinedBounds() instead
 */
function _getObjectBoundsRecursive(item) {
    var bounds = getBounds(item);
    if (bounds) {
        return { L: bounds.left, T: bounds.top, R: bounds.right, B: bounds.bottom };
    }
    return null;
}

/**
 * Create CMYK Magenta color (for brand guideline styling)
 */
function _cmykMagenta() {
    var c = new CMYKColor();
    c.cyan = 0;
    c.magenta = 100;
    c.yellow = 0;
    c.black = 0;
    return c;
}

/**
 * Add dashed rectangle (brand guideline style)
 */
function _addDashedRect(layer, left, top, right, bottom, strokeW, color) {
    var w = right - left;
    var h = top - bottom;
    var r = layer.pathItems.rectangle(top, left, w, h);
    r.stroked = true;
    r.strokeWidth = strokeW;
    r.strokeColor = color;
    r.filled = false;
    
    // Create dashed stroke (Illustrator dash pattern)
    try {
        r.strokeDashes = [4, 2]; // 4pt dash, 2pt gap
        r.strokeDashCap = StrokeCap.ROUNDENDCAP;
    } catch(e) {}
    
    return r;
}

/**
 * Add "X" label text (brand guideline annotation)
 */
function _addXLabel(layer, x, y, sizePt, color) {
    var tf = layer.textFrames.pointText([x, y]);
    tf.contents = 'X';
    try {
        tf.textRange.characterAttributes.size = sizePt;
        tf.textRange.characterAttributes.fillColor = color;
        try {
            tf.textRange.characterAttributes.font = app.textFonts.getByName('ArialMT');
        } catch(e) {
            // Use default font if Arial not available
        }
        tf.textRange.characterAttributes.strokeWeight = 0;
    } catch(e) {
        // Fallback if styling fails
        try {
            tf.textRange.characterAttributes.size = sizePt;
            tf.textRange.characterAttributes.fillColor = color;
        } catch(e2) {}
    }
    return tf;
}


/**
 * Akrivi Clearspace Generator
 * Creates professional brand-guideline-ready clearspace diagrams
 * 
 * @param xSourceIndex - Index of X source object in selection (-1 for fallback)
 * @param xMode - "width", "height", or "both"
 * @param multiplier - Clearspace multiplier (default 1X)
 * @param updateMode - true to update existing, false to create new
 */
/**
 * Akrivi Clearspace Grid Generator (3×3) - Integrated JSX Engine
 * Creates brand-guideline-ready clearspace diagrams with Logo Mark, Logo Type, and Exclusion Zone components
 * Based on the Akrivi methodology: strict 3×3 grid with X-unit visualization
 */
function helpersGenerateAkriviClearspaceGrid(logoMarkIndex, logoTypeIndex, xSourceIndex, logoMarkName, logoTypeName, xSourceName, lockupMode, xDefinition, lockupGapScale, clearspaceScale, updateMode) {
    try {
        if (!app.documents.length) return 'No document open.';
        var doc = app.activeDocument;
        
        // Validate X Source is required
        if (xSourceIndex < 0) return 'ERROR: Exclusion Zone component is required.';
        
        var lockupModeStr = String(lockupMode || 'horizontal');
        var xDefStr = String(xDefinition || 'width');
        var lockupGapPercent = Number(lockupGapScale); if (!lockupGapPercent || lockupGapPercent <= 0) lockupGapPercent = 100;
        var clearspacePercent = Number(clearspaceScale); if (!clearspacePercent || clearspacePercent <= 0) clearspacePercent = 100;
        
        // Get component objects by name (more reliable than indices)
        // Search all layers for items matching stored names
        var logoMark = null;
        var logoType = null;
        var xSource = null;
        
        var logoMarkNameStr = String(logoMarkName || '');
        var logoTypeNameStr = String(logoTypeName || '');
        var xSourceNameStr = String(xSourceName || '');
        
        // Search function to find item by name
        function findItemByName(doc, itemName) {
            if (!itemName || itemName === '') return null;
            for (var l = 0; l < doc.layers.length; l++) {
                var layer = doc.layers[l];
                try {
                    for (var i = 0; i < layer.pageItems.length; i++) {
                        try {
                            if (layer.pageItems[i].name === itemName) {
                                return layer.pageItems[i];
                            }
                        } catch(e) {}
                    }
                } catch(e) {}
            }
            return null;
        }
        
        // Try to find components by name first
        if (logoMarkNameStr && logoMarkNameStr !== '') {
            logoMark = findItemByName(doc, logoMarkNameStr);
        }
        if (logoTypeNameStr && logoTypeNameStr !== '') {
            logoType = findItemByName(doc, logoTypeNameStr);
        }
        if (xSourceNameStr && xSourceNameStr !== '') {
            xSource = findItemByName(doc, xSourceNameStr);
        }
        
        // Fallback: Try selection-based lookup if name search failed
        if (!xSource && doc.selection && doc.selection.length > 0) {
            if (xSourceIndex >= 0 && xSourceIndex < doc.selection.length) {
                try { xSource = doc.selection[xSourceIndex]; } catch(e) {}
            }
        }
        if (!logoMark && doc.selection && doc.selection.length > 0) {
            if (logoMarkIndex >= 0 && logoMarkIndex < doc.selection.length) {
                try { logoMark = doc.selection[logoMarkIndex]; } catch(e) {}
            }
        }
        if (!logoType && doc.selection && doc.selection.length > 0) {
            if (logoTypeIndex >= 0 && logoTypeIndex < doc.selection.length) {
                try { logoType = doc.selection[logoTypeIndex]; } catch(e) {}
            }
        }
        
        // Validate X Source is required
        if (!xSource) return 'ERROR: Exclusion Zone component not found. Please select it again.';
        
        // Create a duplicate named "X-REF"; do NOT rename or modify the original
        var xRefCopy;
        try {
            xRefCopy = xSource.duplicate();
            xRefCopy.name = 'X-REF';
        } catch (e) {
            return 'ERROR: Could not create X-REF copy: ' + (e.toString ? e.toString() : e);
        }
        
        // Utility functions (from JSX script)
        function safeBounds(item) {
            try { 
                if (item.visibleBounds) {
                    var vb = item.visibleBounds;
                    return [vb[0], vb[1], vb[2], vb[3]]; // [left, top, right, bottom]
                }
            } catch (e1) {}
            try { 
                var gb = item.geometricBounds;
                return [gb[0], gb[1], gb[2], gb[3]];
            } catch (e2) {}
            return null;
        }
        
        function boundsUnion(a, b) {
            if (!a) return b;
            if (!b) return a;
            return [
                Math.min(a[0], b[0]),
                Math.max(a[1], b[1]),
                Math.max(a[2], b[2]),
                Math.min(a[3], b[3])
            ];
        }
        
        function bw(b) { return b[2] - b[0]; } // width
        function bh(b) { return b[1] - b[3]; } // height
        
        // Get X value from item (maps xDefinition to JSX script mode)
        function getXValueFromItem(item, mode) {
            var b = safeBounds(item);
            if (!b) return null;
            
            var w = Math.max(0.0001, bw(b));
            var h = Math.max(0.0001, bh(b));
            
            if (mode === 'width' || mode === '1') return w;
            if (mode === 'height' || mode === '2') return h;
            
            // AUTO: farthest extent (square-safe) - matches JSX script
            return Math.max(w, h);
        }
        
        // -------------------------
        // 1) Compute logo bounds L (union of lockup components)
        // -------------------------
        var logoBounds = null;
        var lockupItems = [];
        if (logoMark) lockupItems.push(logoMark);
        if (logoType) lockupItems.push(logoType);
        
        if (lockupItems.length === 0) {
            return 'ERROR: At least Logo Mark or Logo Type must be selected.';
        }
        
        // Build lockup bounds (union of selected components)
        for (var iSel = 0; iSel < lockupItems.length; iSel++) {
            var bSel = safeBounds(lockupItems[iSel]);
            if (bSel) logoBounds = boundsUnion(logoBounds, bSel);
        }
        
        // If both Mark and Type exist, calculate lockup with gap
        if (logoMark && logoType && lockupModeStr) {
            var markB = safeBounds(logoMark);
            var typeB = safeBounds(logoType);
            
            if (markB && typeB) {
                // Calculate gap based on X-unit and lockup gap scale (use X-REF copy)
                var tempX = getXValueFromItem(xRefCopy, xDefStr);
                if (!tempX || tempX <= 0) tempX = 0.1 * bh(logoBounds);
                var gap = (tempX * lockupGapPercent) / 100;
                
                if (lockupModeStr === 'horizontal') {
                    // Horizontal: Mark left, Type right
                    logoBounds = [
                        markB[0], // left
                        Math.max(markB[1], typeB[1]), // top
                        markB[2] + gap + bw(typeB), // right
                        Math.min(markB[3], typeB[3]) // bottom
                    ];
        } else {
                    // Vertical: Mark top, Type bottom
                    logoBounds = [
                        Math.min(markB[0], typeB[0]), // left
                        markB[1], // top
                        Math.max(markB[2], typeB[2]), // right
                        markB[3] - gap - bh(typeB) // bottom
                    ];
                }
            }
        }
        
        if (!logoBounds) {
            return 'ERROR: Could not calculate logo bounds.';
        }
        
        var Lw = bw(logoBounds);
        var Lh = bh(logoBounds);
        if (!isFinite(Lw) || !isFinite(Lh) || Lw <= 0 || Lh <= 0) {
            return 'ERROR: Invalid logo bounds.';
        }
        
        // -------------------------
        // 2) Compute X (from X-REF copy or fallback)
        // -------------------------
        var X;
        var xMode = xDefStr; // 'width', 'height', or 'auto'
        
        var xVal = getXValueFromItem(xRefCopy, xMode);
        if (xVal !== null && xVal > 0) {
            X = xVal;
        } else {
            X = 0.1 * Lh; // fallback
        }
        
        if (!isFinite(X) || X <= 0) X = 10;
        X = Math.max(X, 0.5);
        
        // Apply clearspace scaling
        X = (X * clearspacePercent) / 100;
        X = Math.max(X, 0.5);
        
        // For 3×3 grid, we use uniform X (square-safe)
        var clearspaceX = X;
        var clearspaceY = X;
        
        // -------------------------
        // 3) Build 3×3 grid geometry (based on logo bounds)
        // -------------------------
        // Grid structure: [X] [Lw] [X] columns, [X] [Lh] [X] rows
        // Total: Lw + 2X width, Lh + 2X height
        
        var gridLeft   = logoBounds[0] - X;
        var gridTop    = logoBounds[1] + X;
        var gridRight  = logoBounds[2] + X;
        var gridBottom = logoBounds[3] - X;
        
        // Grid line positions (inner corners of center cell)
        var x1 = gridLeft + X;   // Left edge of center column
        var x2 = x1 + Lw;        // Right edge of center column
        
        var y1 = gridTop - X;    // Top edge of center row
        var y2 = y1 - Lh;        // Bottom edge of center row
        
        // Corner cell centers (for circle fallback)
        function centerOfCell(left, top, right, bottom) {
            return [(left + right) / 2, (top + bottom) / 2];
        }
        
        var TLc = centerOfCell(gridLeft, gridTop, x1, y1);
        var TRc = centerOfCell(x2, gridTop, gridRight, y1);
        var BLc = centerOfCell(gridLeft, y2, x1, gridBottom);
        var BRc = centerOfCell(x2, y2, gridRight, gridBottom);
        
        // Inner corner intersection points (for FLUSH alignment of clones)
        var innerTL = { x: x1, y: y1 };
        var innerTR = { x: x2, y: y1 };
        var innerBL = { x: x1, y: y2 };
        var innerBR = { x: x2, y: y2 };
        
        // -------------------------
        // 4) Draw output
        // -------------------------
        // Create or find dedicated layer
        var clearspaceLayer = null;
        for (var l = 0; l < doc.layers.length; l++) {
            if (doc.layers[l].name === 'Helpersub_Clearspace') {
                clearspaceLayer = doc.layers[l];
                break;
            }
        }
        
        if (!clearspaceLayer) {
            clearspaceLayer = doc.layers.add();
            clearspaceLayer.name = 'Helpersub_Clearspace';
            try {
                clearspaceLayer.printable = false;
                clearspaceLayer.locked = false;
                clearspaceLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
            } catch(e) {}
        } else {
            // Clear existing content on regeneration (always clean slate)
            // Remove all pageItems from the layer (including groups)
            for (var p = clearspaceLayer.pageItems.length - 1; p >= 0; p--) {
                try {
                    var item = clearspaceLayer.pageItems[p];
                    // Remove our clearspace group or any other items
                    if (item.typename === 'GroupItem' && (item.name === 'Clearspace_Grid_Output' || item.name === 'Akrivi_ClearspaceGrid')) {
                        item.remove();
                    } else {
                        // Remove any other items (cleanup)
                        item.remove();
                    }
                } catch(e) {}
            }
        }
        
        // Ensure layer is unlocked and visible
        try {
            clearspaceLayer.locked = false;
            clearspaceLayer.visible = true;
        } catch(e) {}
        
        // Clear old output
        try {
            for (var i = clearspaceLayer.groupItems.length - 1; i >= 0; i--) {
                if (clearspaceLayer.groupItems[i].name === "Clearspace_Grid_Output") {
                    clearspaceLayer.groupItems[i].remove();
                }
            }
        } catch (e) {}
        
        var group = clearspaceLayer.groupItems.add();
        group.name = "Clearspace_Grid_Output";
        
        // Style (matching JSX script)
        function makeRGB(r, g, b) {
            var c = new RGBColor();
            c.red = r; c.green = g; c.blue = b;
            return c;
        }
        
        var GRID_STROKE   = makeRGB(120, 180, 210);
        var CIRCLE_STROKE = makeRGB(120, 180, 210);
        var CIRCLE_FILL   = makeRGB(210, 235, 245);
        var strokeW = 0.5;
        var cornerOpacity = 18; // From JSX script default
        
        // Helper functions for drawing
        function addLine(parent, x1, y1, x2, y2, strokeColor, strokeWidth) {
            var p = parent.pathItems.add();
            p.stroked = true;
            p.filled = false;
            p.strokeColor = strokeColor;
            p.strokeWidth = strokeWidth;
            p.setEntirePath([[x1, y1], [x2, y2]]);
            return p;
        }
        
        function addCircle(parent, cx, cy, diameter, strokeColor, strokeWidth, fillColor) {
            var r = diameter / 2;
            var top = cy + r;
            var left = cx - r;
            var el = parent.pathItems.ellipse(top, left, diameter, diameter);
            el.stroked = true;
            el.strokeColor = strokeColor;
            el.strokeWidth = strokeWidth;
            el.filled = true;
            el.fillColor = fillColor;
            return el;
        }
        
        function addRect(parent, left, top, width, height, strokeColor, strokeWidth) {
            var r = parent.pathItems.rectangle(top, left, width, height);
            r.stroked = true;
            r.filled = false;
            r.strokeColor = strokeColor;
            r.strokeWidth = strokeWidth;
            return r;
        }
        
        // Outer rect + grid lines
        addRect(group, gridLeft, gridTop, (gridRight - gridLeft), (gridTop - gridBottom), GRID_STROKE, strokeW);
        
        // Vertical lines
        addLine(group, gridLeft,  gridTop, gridLeft,  gridBottom, GRID_STROKE, strokeW);
        addLine(group, x1,        gridTop, x1,        gridBottom, GRID_STROKE, strokeW);
        addLine(group, x2,        gridTop, x2,        gridBottom, GRID_STROKE, strokeW);
        addLine(group, gridRight, gridTop, gridRight, gridBottom, GRID_STROKE, strokeW);
        
        // Horizontal lines
        addLine(group, gridLeft, gridTop,    gridRight, gridTop,    GRID_STROKE, strokeW);
        addLine(group, gridLeft, y1,         gridRight, y1,         GRID_STROKE, strokeW);
        addLine(group, gridLeft, y2,         gridRight, y2,         GRID_STROKE, strokeW);
        addLine(group, gridLeft, gridBottom, gridRight, gridBottom, GRID_STROKE, strokeW);
        
        // -------------------------
        // 5) Corner unit visuals (X clones or circles)
        // -------------------------
        if (xSource) {
            // Place clones FLUSH to inner corner intersections
            function setOpacity(item, value) {
                try { item.opacity = value; } catch (e) {}
            }
            
            function rotateAboutCenter(item, deg) {
                try {
                    item.rotate(deg, true, true, true, true, Transformation.CENTER);
    } catch (e) {
                    try { item.rotate(deg); } catch (e2) {}
                }
            }
            
            // Fit INSIDE a square X×X (no protruding tips)
            function fitItemToX(item, X) {
                var b = safeBounds(item);
                if (!b) return;
                var w = Math.max(0.0001, bw(b));
                var h = Math.max(0.0001, bh(b));
                var current = Math.max(w, h);
                if (!isFinite(current) || current <= 0) return;
                var scale = X / current;
                var pct = scale * 100;
                try {
                    item.resize(pct, pct, true, true, true, true, pct, Transformation.CENTER);
                } catch (e) {
                    try { item.resize(pct, pct); } catch (e2) {}
                }
            }
            
            // Align a specific corner of ITEM bounds to a point
            function alignItemCornerToPoint(item, corner, px, py) {
                var b = safeBounds(item);
                if (!b) return;
                var left = b[0], top = b[1], right = b[2], bottom = b[3];
                var ax, ay;
                if (corner === "TL") { ax = left;  ay = top; }
                if (corner === "TR") { ax = right; ay = top; }
                if (corner === "BL") { ax = left;  ay = bottom; }
                if (corner === "BR") { ax = right; ay = bottom; }
                item.translate(px - ax, py - ay);
            }
            
            function cloneToInnerCorner(innerX, innerY, rotationDeg, alignCorner) {
                var clone = xSource.duplicate(group, ElementPlacement.PLACEATEND);
                fitItemToX(clone, X);
                rotateAboutCenter(clone, rotationDeg);
                fitItemToX(clone, X); // Re-fit after rotation
                alignItemCornerToPoint(clone, alignCorner, innerX, innerY);
                setOpacity(clone, cornerOpacity);
                return clone;
            }
            
            // Alternating 0/90, FLUSH to inner corners (matching JSX script)
            cloneToInnerCorner(innerTL.x, innerTL.y, 0,  "BR"); // TL cell
            cloneToInnerCorner(innerTR.x, innerTR.y, 90, "BL"); // TR cell
            cloneToInnerCorner(innerBR.x, innerBR.y, 0,  "TL"); // BR cell
            cloneToInnerCorner(innerBL.x, innerBL.y, 90, "TR"); // BL cell
        } else {
            // Fallback: corner circles centered in corner cells
            addCircle(group, TLc[0], TLc[1], X, CIRCLE_STROKE, strokeW, CIRCLE_FILL);
            addCircle(group, TRc[0], TRc[1], X, CIRCLE_STROKE, strokeW, CIRCLE_FILL);
            addCircle(group, BLc[0], BLc[1], X, CIRCLE_STROKE, strokeW, CIRCLE_FILL);
            addCircle(group, BRc[0], BRc[1], X, CIRCLE_STROKE, strokeW, CIRCLE_FILL);
        }
        
        try { group.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eG) {}
        try { clearspaceLayer.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eL) {}
        
        // Build success message
        var srcName = "(fallback 10% of logo height)";
        if (xRefCopy) {
            try { 
                srcName = (xRefCopy.name && xRefCopy.name.length) ? xRefCopy.name : xRefCopy.typename; 
            } catch (eN) { 
                srcName = "X-REF"; 
            }
        }
        
        var modeStr = (xMode === 'width' || xMode === '1') ? 'WIDTH' : ((xMode === 'height' || xMode === '2') ? 'HEIGHT' : 'AUTO (square-safe)');
        return 'Clearspace grid generated.\n' +
               'Layer: Helpersub_Clearspace (non-printing)\n' +
               'Mode: ' + modeStr + '\n' +
               'X = ' + (Math.round(X * 100) / 100) + ' px\n' +
               'X Source: ' + srcName + '\n' +
               (xRefCopy ? ('Corner clones opacity: ' + cornerOpacity + '% (flush to inner corners)') : 'Corner units: circles');
    } catch (e) {
        return 'ERROR: ' + e.toString() + ' (Line: ' + (e.line || 'unknown') + ')';
    }
}
// Explicitly attach to $.global for CEP access
try { if (typeof $.global !== 'undefined') $.global.helpersGenerateAkriviClearspaceGrid = helpersGenerateAkriviClearspaceGrid; } catch(e) {}

// Keep old function for backward compatibility (deprecated)
function helpersGenerateAkriviClearspace(xSourceIndex, xMode, multiplier, updateMode) {
    // Legacy wrapper - redirects to new grid function with defaults
    return helpersGenerateAkriviClearspaceGrid(-1, -1, xSourceIndex, '', '', '', 'horizontal', xMode, 100, multiplier * 100, updateMode);
}
try { if (typeof $.global !== 'undefined') $.global.helpersGenerateAkriviClearspace = helpersGenerateAkriviClearspace; } catch(e) {}

/**
 * Simple JSON.stringify implementation for ExtendScript
 * ExtendScript doesn't have native JSON support
 */
function _jsonStringify(obj) {
    if (obj === null) return 'null';
    if (obj === undefined) return 'undefined';
    if (typeof obj === 'string') {
        // Escape special characters
        var str = obj.replace(/\\/g, '\\\\')
                     .replace(/"/g, '\\"')
                     .replace(/\n/g, '\\n')
                     .replace(/\r/g, '\\r')
                     .replace(/\t/g, '\\t');
        return '"' + str + '"';
    }
    if (typeof obj === 'number') return String(obj);
    if (typeof obj === 'boolean') return obj ? 'true' : 'false';
    if (obj instanceof Array) {
        var parts = [];
        for (var i = 0; i < obj.length; i++) {
            parts.push(_jsonStringify(obj[i]));
        }
        return '[' + parts.join(',') + ']';
    }
    if (typeof obj === 'object') {
        var parts = [];
        for (var key in obj) {
            try {
                // Check if property exists and is not a function
                if (typeof obj[key] === 'function') continue;
                var value = obj[key];
                if (value !== undefined) {
                    parts.push('"' + String(key).replace(/"/g, '\\"') + '":' + _jsonStringify(value));
                }
            } catch(e) {
                // Skip properties that can't be accessed
                continue;
            }
        }
        return '{' + parts.join(',') + '}';
    }
    return 'null';
}

/**
 * Get selection info for minimal Clearspace UI (returns JSON string).
 * Reads current active selection only (app.activeDocument.selection) — no cached or stored reference.
 * Returns selection items with bounds info and recommended index.
 */
function helpersGetSelectionInfoMinimal() {
    try {
        if (!app.documents.length) {
            return _jsonStringify({ hasSelection: false, items: [], recommendedIndex: -1 });
        }
        
        var doc = app.activeDocument;
        // Current active selection only — never use cached/stored selection
        if (!doc.selection || doc.selection.length === 0) {
            return _jsonStringify({ hasSelection: false, items: [], recommendedIndex: -1 });
        }
        
        var items = [];
        var recommendedIndex = doc.selection.length - 1; // Prefer last selected
        
        for (var i = 0; i < doc.selection.length; i++) {
            try {
                var item = doc.selection[i];
                if (!item) continue;
                
                var bounds = null;
                try {
                    if (item.visibleBounds) {
                        bounds = item.visibleBounds;
                    } else if (item.geometricBounds) {
                        bounds = item.geometricBounds;
                    }
                } catch(e) {}
                
                var typename = '';
                var name = '';
                var widthPt = 0;
                var heightPt = 0;
                
                try {
                    typename = item.typename || 'Unknown';
                } catch(e) {
                    typename = 'Unknown';
                }
                
                try {
                    name = item.name || '';
                } catch(e) {
                    name = '';
                }
                
                if (bounds && bounds.length >= 4) {
                    widthPt = Math.abs(bounds[2] - bounds[0]);
                    heightPt = Math.abs(bounds[1] - bounds[3]);
                }
                
                items.push({
                    index: i,
                    typename: typename,
                    name: name,
                    widthPt: Math.round(widthPt * 100) / 100,
                    heightPt: Math.round(heightPt * 100) / 100
                });
            } catch(e) {
                // Skip invalid items
            }
        }
        
        return _jsonStringify({
            hasSelection: true,
            items: items,
            recommendedIndex: recommendedIndex
        });
    } catch (e) {
        return _jsonStringify({ 
            hasSelection: false, 
            items: [], 
            recommendedIndex: -1,
            error: e.toString()
        });
    }
}
try { if (typeof $.global !== 'undefined') $.global.helpersGetSelectionInfoMinimal = helpersGetSelectionInfoMinimal; } catch(e) {}

/**
 * Force delete all page items named "X-REF" in the document (all layers).
 * Used before creating a new X-REF and when cleaning clearspace so old X-REF data is completely removed.
 * Returns a string message for UI.
 */
function helpersForceDeleteAllXRef() {
    try {
        if (!app.documents.length) return "No document open.";
        var doc = app.activeDocument;
        var count = 0;
        for (var L = 0; L < doc.layers.length; L++) {
            var layer = doc.layers[L];
            var items = layer.pageItems;
            // Iterate backwards so removal does not shift indices
            for (var i = items.length - 1; i >= 0; i--) {
                try {
                    if (items[i].name === "X-REF") {
                        items[i].remove();
                        count++;
                    }
                } catch (e) {}
            }
        }
        return count > 0 ? "Removed " + count + " X-REF object(s)." : "No X-REF objects found.";
    } catch (e) {
        return "ERROR: " + (e.toString ? e.toString() : e);
    }
}
try { if (typeof $.global !== 'undefined') $.global.helpersForceDeleteAllXRef = helpersForceDeleteAllXRef; } catch(e) {}

/**
 * Create a duplicate of the selected item, name it "X-REF", and return its info.
 * Does NOT rename or move the original; only the copy is named "X-REF".
 * Gets selection first, duplicates it, then removes old X-REFs so we never delete the selected item before copying.
 */
function helpersSetXRef(index) {
    try {
        if (!app.documents.length) {
            return _jsonStringify({ success: false, error: 'No document open' });
        }
        
        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            return _jsonStringify({ success: false, error: 'No selection' });
        }
        
        if (index < 0 || index >= doc.selection.length) {
            return _jsonStringify({ success: false, error: 'Invalid index' });
        }
        
        // Get the selected item FIRST (before any deletion) so selection is still valid
        var item = doc.selection[index];
        if (!item) {
            return _jsonStringify({ success: false, error: 'Selection item not found (selection may have changed)' });
        }
        if (typeof item.duplicate !== 'function') {
            return _jsonStringify({ success: false, error: 'Selected item cannot be duplicated' });
        }
        
        var copy;
        try {
            copy = item.duplicate();
        } catch (e) {
            return _jsonStringify({ success: false, error: 'Could not duplicate: ' + (e && e.toString ? e.toString() : String(e)) });
        }
        
        // Now safe to remove any existing X-REF (we already have our copy)
        helpersForceDeleteAllXRef();
        
        try {
            copy.name = 'X-REF';
        } catch (e) {
            try { copy.remove(); } catch (e2) {}
            return _jsonStringify({ success: false, error: 'Could not name copy: ' + e.toString() });
        }
        
        // Get bounds of the copy (same as original)
        var bounds = null;
        try {
            if (copy.visibleBounds) {
                bounds = copy.visibleBounds;
            } else if (copy.geometricBounds) {
                bounds = copy.geometricBounds;
            }
        } catch(e) {}
        
        var typename = '';
        try {
            typename = copy.typename || 'Unknown';
        } catch(e) {
            typename = 'Unknown';
        }
        
        var widthPt = 0;
        var heightPt = 0;
        if (bounds && bounds.length >= 4) {
            widthPt = Math.abs(bounds[2] - bounds[0]);
            heightPt = Math.abs(bounds[1] - bounds[3]);
        }
        
        return _jsonStringify({
            success: true,
            typename: typename,
            name: 'X-REF',
            widthPt: Math.round(widthPt * 100) / 100,
            heightPt: Math.round(heightPt * 100) / 100
        });
    } catch (e) {
        return _jsonStringify({ success: false, error: (e && e.toString ? e.toString() : String(e)) });
    }
}
try { if (typeof $.global !== 'undefined') $.global.helpersSetXRef = helpersSetXRef; } catch(e) {}

/**
 * Run Clearspace generator (adapted from standalone JSX)
 * Accepts parameters instead of prompts
 * @param xMode - "1"=width, "2"=height, "3"=auto
 * @param strokeW - Stroke width in points
 * @param cornerOpacity - Opacity for corner clones (0-100)
 * @param clearspacePercent - Clearspace scaling percentage (default 100)
 * @param updateMode - true to update existing output, false to create new
 */
function helpersRunClearspace(xMode, strokeW, cornerOpacity, clearspacePercent, updateMode, colorR, colorG, colorB) {
    try {
        if (app.documents.length === 0) {
            return 'ERROR: No document open.';
        }

        var doc = app.activeDocument;
        var sel = doc.selection;

        if (!sel || sel.length === 0) {
            return 'ERROR: Select the logo (and optionally an X source object).';
        }

        // Use provided parameters or defaults
        var mode = String(xMode || '3'); // '1'=width, '2'=height, '3'=auto
        var strokeWidth = parseFloat(strokeW);
        if (isNaN(strokeWidth) || strokeWidth <= 0) strokeWidth = 0.5;
        
        var opacity = parseFloat(cornerOpacity);
        if (isNaN(opacity)) opacity = 18;
        opacity = Math.max(0, Math.min(100, opacity));
        
        var clearspacePercentValue = parseFloat(clearspacePercent);
        if (isNaN(clearspacePercentValue) || clearspacePercentValue <= 0) clearspacePercentValue = 100;

        // Utilities (from JSX script)
        function safeBounds(item) {
            try { if (item.visibleBounds) return item.visibleBounds; } catch (e1) {}
            try { return item.geometricBounds; } catch (e2) {}
            return null;
        }

        function boundsUnion(a, b) {
            if (!a) return b;
            if (!b) return a;
            return [
                Math.min(a[0], b[0]),
                Math.max(a[1], b[1]),
                Math.max(a[2], b[2]),
                Math.min(a[3], b[3])
            ];
        }

        function bw(b) { return b[2] - b[0]; }
        function bh(b) { return b[1] - b[3]; }

        function makeRGB(r, g, b) {
            var c = new RGBColor();
            c.red = r; c.green = g; c.blue = b;
            return c;
        }

        function ensureLayer(name) {
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === name) return doc.layers[i];
            }
            var L = doc.layers.add();
            L.name = name;
            return L;
        }

        function clearOldOutput(layer) {
            try {
                for (var i = layer.groupItems.length - 1; i >= 0; i--) {
                    if (layer.groupItems[i].name === "Clearspace_Grid_Output") {
                        layer.groupItems[i].remove();
                    }
                }
            } catch (e) {}
        }

        function addLine(parent, x1, y1, x2, y2, strokeColor, strokeWidth) {
            var p = parent.pathItems.add();
            p.stroked = true;
            p.filled = false;
            p.strokeColor = strokeColor;
            p.strokeWidth = strokeWidth;
            p.setEntirePath([[x1, y1], [x2, y2]]);
            return p;
        }

        function addCircle(parent, cx, cy, diameter, strokeColor, strokeWidth, fillColor) {
            var r = diameter / 2;
            var top = cy + r;
            var left = cx - r;
            var el = parent.pathItems.ellipse(top, left, diameter, diameter);
            el.stroked = true;
            el.strokeColor = strokeColor;
            el.strokeWidth = strokeWidth;
            el.filled = true;
            el.fillColor = fillColor;
            return el;
        }

        function addRect(parent, left, top, width, height, strokeColor, strokeWidth) {
            var r = parent.pathItems.rectangle(top, left, width, height);
            r.stroked = true;
            r.filled = false;
            r.strokeColor = strokeColor;
            r.strokeWidth = strokeWidth;
            return r;
        }

        function getXValueFromItem(item, mode) {
            var b = safeBounds(item);
            if (!b) return null;
            var w = Math.max(0.0001, bw(b));
            var h = Math.max(0.0001, bh(b));
            if (mode === "1") return w;
            if (mode === "2") return h;
            return Math.max(w, h); // AUTO
        }

        function findNamedXRefInAllLayers(doc) {
            // Search all layers EXCEPT the Clear Space output layer to find the original X-REF
            // This ensures we always use the original reference, not previously generated Clear Space objects
            for (var l = 0; l < doc.layers.length; l++) {
                var layer = doc.layers[l];
                // Skip the Clear Space output layer - X-REF should never be there
                if (layer.name === "Helpersub_Clearspace") continue;
                
                try {
                    for (var i = 0; i < layer.pageItems.length; i++) {
                        try {
                            if (layer.pageItems[i].name === "X-REF") {
                                return layer.pageItems[i];
                            }
                        } catch(e) {}
                    }
                } catch(e) {}
            }
            return null;
        }

        function findNamedXRef(selection) {
            // First try to find X-REF in all layers (original reference)
            var xref = findNamedXRefInAllLayers(doc);
            if (xref) return xref;
            
            // Fallback: check current selection (but exclude Clear Space layer items)
            for (var i = 0; i < selection.length; i++) {
                try {
                    if (selection[i].name === "X-REF") {
                        // Verify it's not in the Clear Space output layer
                        try {
                            var parent = selection[i].parent;
                            while (parent) {
                                if (parent.typename === "Layer" && parent.name === "Helpersub_Clearspace") {
                                    // This is in Clear Space layer, skip it
                                    break;
                                }
                                if (parent.typename === "Layer") {
                                    // Found a layer that's not Clear Space, this is valid
                                    return selection[i];
                                }
                                parent = parent.parent;
                            }
                        } catch(e) {
                            // If we can't check parent, use it anyway (fallback)
                            return selection[i];
                        }
                    }
                } catch (e) {}
            }
            return null;
        }

        function pickSmartXSource(selection, mode) {
            // Always prioritize finding the original X-REF by name across all layers
            var xref = findNamedXRef(selection);
            if (xref) return { item: xref, reason: 'Found "X-REF" (original reference)' };
            
            // Only use selection-based fallback if no X-REF found by name
            if (selection.length >= 2) {
                var bestItem = null;
                var bestX = null;
                for (var i = 0; i < selection.length; i++) {
                    // Skip items in Clear Space output layer
                    try {
                        var parent = selection[i].parent;
                        var isInClearSpaceLayer = false;
                        while (parent) {
                            if (parent.typename === "Layer" && parent.name === "Helpersub_Clearspace") {
                                isInClearSpaceLayer = true;
                                break;
                            }
                            if (parent.typename === "Layer") break;
                            parent = parent.parent;
                        }
                        if (isInClearSpaceLayer) continue;
                    } catch(e) {}
                    
                    var xVal = getXValueFromItem(selection[i], mode);
                    if (xVal === null) continue;
                    if (bestX === null || xVal < bestX) {
                        bestX = xVal;
                        bestItem = selection[i];
                    }
                }
                if (bestItem) return { item: bestItem, reason: "Picked smallest selected item as X-source" };
            }
            return { item: null, reason: "No explicit X-source (fallback will be used)" };
        }

        function centerOfCell(left, top, right, bottom) {
            return [(left + right) / 2, (top + bottom) / 2];
        }

        function setOpacity(item, value) {
            try { item.opacity = value; } catch (e) {}
        }

        function rotateAboutCenter(item, deg) {
            try {
                item.rotate(deg, true, true, true, true, Transformation.CENTER);
            } catch (e) {
                try { item.rotate(deg); } catch (e2) {}
            }
        }

        function fitItemToX(item, X) {
            var b = safeBounds(item);
            if (!b) return;
            var w = Math.max(0.0001, bw(b));
            var h = Math.max(0.0001, bh(b));
            var current = Math.max(w, h);
            if (!isFinite(current) || current <= 0) return;
            var scale = X / current;
            var pct = scale * 100;
            try {
                item.resize(pct, pct, true, true, true, true, pct, Transformation.CENTER);
            } catch (e) {
                try { item.resize(pct, pct); } catch (e2) {}
            }
        }

        function alignItemCornerToPoint(item, corner, px, py) {
            var b = safeBounds(item);
            if (!b) return;
            var left = b[0], top = b[1], right = b[2], bottom = b[3];
            var ax, ay;
            if (corner === "TL") { ax = left;  ay = top; }
            if (corner === "TR") { ax = right; ay = top; }
            if (corner === "BL") { ax = left;  ay = bottom; }
            if (corner === "BR") { ax = right; ay = bottom; }
            item.translate(px - ax, py - ay);
        }

        // 1) Compute logo bounds L (union of selection)
        var logoBounds = null;
        for (var iSel = 0; iSel < sel.length; iSel++) {
            var bSel = safeBounds(sel[iSel]);
            if (bSel) logoBounds = boundsUnion(logoBounds, bSel);
        }

        if (!logoBounds) {
            return 'ERROR: Could not read selection bounds.';
        }

        var Lw = bw(logoBounds);
        var Lh = bh(logoBounds);
        if (!isFinite(Lw) || !isFinite(Lh) || Lw <= 0 || Lh <= 0) {
            return 'ERROR: Invalid selection bounds.';
        }

        // 2) Compute X (smart source)
        var picked = pickSmartXSource(sel, mode);
        var X;

        if (picked.item) {
            X = getXValueFromItem(picked.item, mode);
        } else {
            X = 0.1 * Lh; // fallback
        }

        if (!isFinite(X) || X <= 0) X = 10;
        X = Math.max(X, 0.5);
        
        // Apply clearspace percentage scaling
        X = (X * clearspacePercentValue) / 100;
        X = Math.max(X, 0.5);

        // 3) Build 3×3 grid geometry
        var gridLeft   = logoBounds[0] - X;
        var gridTop    = logoBounds[1] + X;
        var gridRight  = logoBounds[2] + X;
        var gridBottom = logoBounds[3] - X;

        var x1 = gridLeft + X;
        var x2 = x1 + Lw;

        var y1 = gridTop - X;
        var y2 = y1 - Lh;

        var TLc = centerOfCell(gridLeft, gridTop, x1, y1);
        var TRc = centerOfCell(x2, gridTop, gridRight, y1);
        var BLc = centerOfCell(gridLeft, y2, x1, gridBottom);
        var BRc = centerOfCell(x2, y2, gridRight, gridBottom);

        var innerTL = { x: x1, y: y1 };
        var innerTR = { x: x2, y: y1 };
        var innerBL = { x: x1, y: y2 };
        var innerBR = { x: x2, y: y2 };

        // 4) Draw output
        var layer = ensureLayer("Helpersub_Clearspace");
        layer.visible = true;
        layer.locked = false;
        try { layer.printable = false; } catch (eP) {}
        try { layer.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eZ) {}

        clearOldOutput(layer);

        var group = layer.groupItems.add();
        group.name = "Clearspace_Grid_Output";

        var r = (colorR !== undefined && colorR !== null) ? Number(colorR) : 120;
        var g = (colorG !== undefined && colorG !== null) ? Number(colorG) : 180;
        var b = (colorB !== undefined && colorB !== null) ? Number(colorB) : 210;
        if (!isFinite(r)) r = 120;
        if (!isFinite(g)) g = 180;
        if (!isFinite(b)) b = 210;
        r = Math.max(0, Math.min(255, Math.round(r)));
        g = Math.max(0, Math.min(255, Math.round(g)));
        b = Math.max(0, Math.min(255, Math.round(b)));

        var GRID_STROKE   = makeRGB(r, g, b);
        var CIRCLE_STROKE = makeRGB(r, g, b);
        var CIRCLE_FILL   = makeRGB(Math.min(255, r + 90), Math.min(255, g + 55), Math.min(255, b + 35));

        addRect(group, gridLeft, gridTop, (gridRight - gridLeft), (gridTop - gridBottom), GRID_STROKE, strokeWidth);

        addLine(group, gridLeft,  gridTop, gridLeft,  gridBottom, GRID_STROKE, strokeWidth);
        addLine(group, x1,        gridTop, x1,        gridBottom, GRID_STROKE, strokeWidth);
        addLine(group, x2,        gridTop, x2,        gridBottom, GRID_STROKE, strokeWidth);
        addLine(group, gridRight, gridTop, gridRight, gridBottom, GRID_STROKE, strokeWidth);

        addLine(group, gridLeft, gridTop,    gridRight, gridTop,    GRID_STROKE, strokeWidth);
        addLine(group, gridLeft, y1,         gridRight, y1,         GRID_STROKE, strokeWidth);
        addLine(group, gridLeft, y2,         gridRight, y2,         GRID_STROKE, strokeWidth);
        addLine(group, gridLeft, gridBottom, gridRight, gridBottom, GRID_STROKE, strokeWidth);

        // 5) Corner unit visuals
        if (picked.item) {
            function cloneToInnerCorner(innerX, innerY, rotationDeg, alignCorner) {
                var clone = picked.item.duplicate(group, ElementPlacement.PLACEATEND);
                fitItemToX(clone, X);
                rotateAboutCenter(clone, rotationDeg);
                fitItemToX(clone, X);
                alignItemCornerToPoint(clone, alignCorner, innerX, innerY);
                setOpacity(clone, opacity);
                return clone;
            }

            cloneToInnerCorner(innerTL.x, innerTL.y, 0,  "BR");
            cloneToInnerCorner(innerTR.x, innerTR.y, 90, "BL");
            cloneToInnerCorner(innerBR.x, innerBR.y, 0,  "TL");
            cloneToInnerCorner(innerBL.x, innerBL.y, 90, "TR");
        } else {
            addCircle(group, TLc[0], TLc[1], X, CIRCLE_STROKE, strokeWidth, CIRCLE_FILL);
            addCircle(group, TRc[0], TRc[1], X, CIRCLE_STROKE, strokeWidth, CIRCLE_FILL);
            addCircle(group, BLc[0], BLc[1], X, CIRCLE_STROKE, strokeWidth, CIRCLE_FILL);
            addCircle(group, BRc[0], BRc[1], X, CIRCLE_STROKE, strokeWidth, CIRCLE_FILL);
        }

        try { group.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eG) {}
        try { layer.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (eL) {}

        var modeStr = (mode === "1") ? "WIDTH" : ((mode === "2") ? "HEIGHT" : "AUTO (square-safe)");
        return 'SUCCESS: Clearspace grid generated.\n' +
               'Mode: ' + modeStr + '\n' +
               'X = ' + (Math.round(X * 100) / 100) + ' px';
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}
try { if (typeof $.global !== 'undefined') $.global.helpersRunClearspace = helpersRunClearspace; } catch(e) {}

/**
 * Live update: update ONLY stroke properties (stroke weight, stroke color) on existing Clearspace_Grid_Output.
 * Does NOT modify fill color — preserves original fill of X source objects (e.g. logos).
 * If no such group is found, do nothing. No delete, no re-create.
 */
function helpersUpdateClearspaceVisuals(strokeW, colorR, colorG, colorB) {
    try {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        var layer = _findLayer(doc, 'Helpersub_Clearspace');
        if (!layer) return;
        var group = null;
        for (var i = 0; i < layer.groupItems.length; i++) {
            if (layer.groupItems[i].name === 'Clearspace_Grid_Output') {
                group = layer.groupItems[i];
                break;
            }
        }
        if (!group) return;
        var sw = Number(strokeW);
        if (isNaN(sw) || sw <= 0) sw = 0.5;
        var r = (colorR != null) ? Math.max(0, Math.min(255, Math.round(Number(colorR)))) : 120;
        var g = (colorG != null) ? Math.max(0, Math.min(255, Math.round(Number(colorG)))) : 180;
        var b = (colorB != null) ? Math.max(0, Math.min(255, Math.round(Number(colorB)))) : 210;
        var sc = _rgb(r, g, b);
        function updatePathItem(item) {
            if (!item) return;
            try {
                if (item.stroked) {
                    item.strokeWidth = sw;
                    item.strokeColor = sc;
                }
                /* Do not set item.fillColor — preserve original fill of X source objects (logos, etc.) */
            } catch (e) {}
        }
        function updateRecursive(parent) {
            if (!parent) return;
            var j;
            try {
                if (parent.pathItems && parent.pathItems.length) {
                    for (j = 0; j < parent.pathItems.length; j++) updatePathItem(parent.pathItems[j]);
                }
                if (parent.groupItems && parent.groupItems.length) {
                    for (j = 0; j < parent.groupItems.length; j++) updateRecursive(parent.groupItems[j]);
                }
            } catch (e) {}
        }
        updateRecursive(group);
    } catch (e) {}
}
try { if (typeof $.global !== 'undefined') $.global.helpersUpdateClearspaceVisuals = helpersUpdateClearspaceVisuals; } catch(e) {}

/**
 * Get selection info for UI display (returns JSON string)
 * This function is critical for component selection
 */
function helpersGetSelectionInfo() {
    try {
        if (!app.documents.length) {
            return _jsonStringify({ selectionCount: 0, items: [], error: 'No document open' });
        }
        
        var doc = app.activeDocument;
        if (!doc.selection) {
            return _jsonStringify({ selectionCount: 0, items: [], error: 'No selection object' });
        }
        
        if (doc.selection.length === 0) {
            return _jsonStringify({ selectionCount: 0, items: [], error: 'No objects selected' });
        }
        
        var items = [];
        for (var i = 0; i < doc.selection.length; i++) {
            try {
                var item = doc.selection[i];
                if (!item) continue;
                
                var bounds = _getObjectBoundsRecursive(item);
                var name = '';
                var typename = '';
                
                try {
                    typename = item.typename || 'Unknown';
                } catch(e) {
                    typename = 'Unknown';
                }
                
                try {
                    name = item.name || '';
                    if (!name || name === '') {
                        // Try to get a descriptive name
                        if (typename === 'TextFrame') {
                            try {
                                var textContent = item.contents;
                                if (textContent && textContent.length > 0) {
                                    name = 'Text: ' + textContent.substring(0, 20);
                                } else {
                                    name = 'TextFrame';
                                }
                            } catch(e) {
                                name = 'TextFrame';
                            }
                        } else {
                            name = typename + ' ' + (i + 1);
                        }
                    }
                } catch(e) {
                    name = typename + ' ' + (i + 1);
                }
                
                var width = 0;
                var height = 0;
                if (bounds) {
                    width = bounds.R - bounds.L;
                    height = bounds.T - bounds.B;
                }
                
                items.push({
                    index: i,
                    name: name,
                    typename: typename,
                    width: width,
                    height: height
                });
            } catch(e) {
                // Skip invalid items but continue processing
                try {
                    items.push({
                        index: i,
                        name: 'Error reading object',
                        typename: 'Error',
                        width: 0,
                        height: 0
                    });
                } catch(e2) {
                    // Skip completely if we can't even add error item
                }
            }
        }
        
        return _jsonStringify({
            selectionCount: doc.selection.length,
            items: items,
            success: true
        });
    } catch (e) {
        return _jsonStringify({ 
            error: e.toString(), 
            selectionCount: 0, 
            items: [],
            success: false
        });
    }
}
try { if (typeof $.global !== 'undefined') $.global.helpersGetSelectionInfo = helpersGetSelectionInfo; } catch(e) {}

// Legacy function kept for backward compatibility (deprecated)
function helpersGenerateClearspace(padding, strokeW, mode, showDims, updateMode) {
    // Redirect to Akrivi system with fallback values
    return helpersGenerateAkriviClearspace(-1, 'width', padding / 40, updateMode);
}
try { if (typeof $.global !== 'undefined') $.global.helpersGenerateClearspace = helpersGenerateClearspace; } catch(e) {}

// ===================== Missing Helpers (v8.4.3 hotfix) =====================

function _getSelectionBounds(doc) {
    // returns {L,T,R,B} in document coordinates
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
    var out = { L: ab[0], T: ab[1], R: ab[2], B: ab[3] };

    if (!doc.selection || doc.selection.length === 0) return out;

    var b = doc.selection[0].geometricBounds; // [L,T,R,B]
    var L = b[0], T = b[1], R = b[2], B = b[3];

    for (var i = 1; i < doc.selection.length; i++) {
        try {
            var bb = doc.selection[i].geometricBounds;
            L = Math.min(L, bb[0]);
            T = Math.max(T, bb[1]);
            R = Math.max(R, bb[2]);
            B = Math.min(B, bb[3]);
        } catch (e) {}
    }

    return { L: L, T: T, R: R, B: B };
}

function _addPointText(group, x, y, txt, sizePt, fillColor, opacity) {
    var doc = app.activeDocument;
    var tf = doc.textFrames.pointText([x, y]);
    tf.contents = txt;
    try {
        tf.textRange.characterAttributes.size = sizePt;
        tf.textRange.characterAttributes.fillColor = fillColor;
    } catch (e) {}
    try { tf.opacity = (opacity != null ? opacity : 100); } catch (e2) {}
    // ensure no stroke
    try { tf.textRange.characterAttributes.strokeColor = fillColor; tf.textRange.characterAttributes.strokeWeight = 0; } catch(e3) {}
    try { tf.move(group, ElementPlacement.PLACEATEND); } catch (e4) {}
    return tf;
}

function _drawAnchor(group, anchorXY, sizePt, shape, fillOrStrokeColor, strokeWidth, useFill) {
    var x = anchorXY[0], y = anchorXY[1];
    var s = Math.max(0.1, sizePt);
    var left = x - s / 2;
    var top = y + s / 2;
    var sw = Number(strokeWidth) || 0.5;

    var item;
    if ((shape + '').toLowerCase() === 'square') {
        item = group.pathItems.rectangle(top, left, s, s);
    } else {
        item = group.pathItems.ellipse(top, left, s, s);
    }

    if (useFill) {
        item.filled = true;
        item.fillColor = fillOrStrokeColor;
        item.stroked = false;
    } else {
        item.filled = false;
        item.stroked = true;
        item.strokeColor = fillOrStrokeColor;
        item.strokeWidth = sw;
    }
    try { item.name = 'Anchor'; } catch (e) {}
    try { item.opacity = 100; } catch (e) {}
    return item;
}

function _drawHandle(group, pathPoint, sizePt, shape, strokeColor, strokeWidth, useFill) {
    var ax = pathPoint.anchor[0], ay = pathPoint.anchor[1];
    var ld = pathPoint.leftDirection;
    var rd = pathPoint.rightDirection;
    var sw = Number(strokeWidth) || 0.5;
    var s = Math.max(0.1, sizePt);
    var half = s / 2; // radius / half-size: line must stop here before the dot so stroke never overlaps the shape

    function isSame(a, b) {
        return (Math.abs(a[0] - b[0]) < 0.001) && (Math.abs(a[1] - b[1]) < 0.001);
    }

    // End point for the line: stop at the edge of the dot (half from center), not at center, so stroke never intrudes into square/circle
    function lineEndAtEdge(anchorX, anchorY, centerX, centerY) {
        var dx = centerX - anchorX;
        var dy = centerY - anchorY;
        var dist = Math.sqrt(dx * dx + dy * dy);
        if (dist <= half) return null; // no room for a visible line
        var t = (dist - half) / dist;
        return [anchorX + t * dx, anchorY + t * dy];
    }

    function addSeg(x1, y1, x2, y2) {
        var ln = group.pathItems.add();
        ln.setEntirePath([[x1, y1], [x2, y2]]);
        ln.filled = false;
        ln.stroked = true;
        ln.strokeColor = strokeColor;
        ln.strokeWidth = sw;
        try { ln.name = 'HandleLine'; } catch (e) {}
        return ln;
    }

    function addMarker(x, y) {
        var left = x - s / 2;
        var top = y + s / 2;
        var it;
        if ((shape + '').toLowerCase() === 'square') {
            it = group.pathItems.rectangle(top, left, s, s);
        } else {
            it = group.pathItems.ellipse(top, left, s, s);
        }
        if (useFill) {
            it.filled = true;
            it.fillColor = strokeColor;
            it.stroked = false;
        } else {
            it.filled = false;
            it.stroked = true;
            it.strokeColor = strokeColor;
            it.strokeWidth = sw;
        }
        try { it.name = 'HandleDot'; } catch (e) {}
        return it;
    }

    if (ld && !isSame(ld, pathPoint.anchor)) {
        var endL = lineEndAtEdge(ax, ay, ld[0], ld[1]);
        if (endL) addSeg(ax, ay, endL[0], endL[1]);
        addMarker(ld[0], ld[1]);
    }
    if (rd && !isSame(rd, pathPoint.anchor)) {
        var endR = lineEndAtEdge(ax, ay, rd[0], rd[1]);
        if (endR) addSeg(ax, ay, endR[0], endR[1]);
        addMarker(rd[0], rd[1]);
    }
}

function _outlineItem(item, strokeColor, strokeWidth) {
    // Non-destructive: duplicate on top and outline it (stroke only)
    var doc = app.activeDocument;
    var dup;

    try {
        dup = item.duplicate();
    } catch (e) {
        return null;
    }

    // If text: convert duplicate to outlines
    try {
        if (dup.typename === 'TextFrame') {
            var outlined = dup.createOutline();
            try { dup.remove(); } catch(e2) {}
            dup = outlined;
        }
    } catch (e3) {}

    // Apply stroke-only styling recursively
    function stylize(pi) {
        try {
            pi.filled = false;
            pi.stroked = true;
            pi.strokeColor = strokeColor;
            pi.strokeWidth = strokeWidth;
        } catch (e) {}
    }

    function walk(node) {
        if (!node) return;
        if (node.typename === 'PathItem') {
            stylize(node);
        } else if (node.typename === 'CompoundPathItem') {
            for (var i = 0; i < node.pathItems.length; i++) stylize(node.pathItems[i]);
        } else if (node.typename === 'GroupItem') {
            for (var j = 0; j < node.pageItems.length; j++) walk(node.pageItems[j]);
        }
    }

    walk(dup);

    // Bring to front so user sees it, but it will not hide original (no fill)
    try { dup.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (e4) {}

    return dup;
}

function _drawGridlines(group, bounds, strokeColor, strokeWidth, gridDir) {
    // Clean structural gridlines: extend key X/Y coordinates to artboard edges
    var doc = app.activeDocument;
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
    var aL = ab[0], aT = ab[1], aR = ab[2], aB = ab[3];

    // Collect anchor coordinates from selection
    var xs = [], ys = [];
    if (doc.selection && doc.selection.length) {
        for (var i = 0; i < doc.selection.length; i++) {
            var paths = _collectPathItems(doc.selection[i]);
            for (var p = 0; p < paths.length; p++) {
                var pi = paths[p];
                for (var k = 0; k < pi.pathPoints.length; k++) {
                    var a = pi.pathPoints[k].anchor;
                    xs.push(a[0]);
                    ys.push(a[1]);
                }
            }
        }
    } else {
        xs.push(bounds.L, bounds.R);
        ys.push(bounds.T, bounds.B);
    }

    function sortNum(a, b) { return a - b; }
    function uniqMerge(arr, tol) {
        arr.sort(sortNum);
        var out = [];
        for (var i = 0; i < arr.length; i++) {
            var v = arr[i];
            if (!out.length || Math.abs(v - out[out.length - 1]) > tol) out.push(v);
        }
        return out;
    }

    xs = uniqMerge(xs, 0.5);
    ys = uniqMerge(ys, 0.5);

    function addLine(x1, y1, x2, y2) {
        var ln = group.pathItems.add();
        ln.setEntirePath([[x1, y1], [x2, y2]]);
        ln.filled = false;
        ln.stroked = true;
        ln.strokeColor = strokeColor;
        ln.strokeWidth = strokeWidth;
        try { ln.name = 'Gridline'; } catch (e) {}
        return ln;
    }

    var dir = (gridDir || 'both').toLowerCase();
    var doV = (dir === 'both' || dir === 'vertical' || dir === 'v');
    var doH = (dir === 'both' || dir === 'horizontal' || dir === 'h');

    // Clamp coordinates to artboard so H/V lines stay strictly within bounds
    if (doV) {
        for (var xi = 0; xi < xs.length; xi++) {
            var xClamp = Math.max(aL, Math.min(aR, xs[xi]));
            addLine(xClamp, aT, xClamp, aB);
        }
    }

    if (doH) {
        for (var yi = 0; yi < ys.length; yi++) {
            var yClamp = Math.max(aB, Math.min(aT, ys[yi]));
            addLine(aL, yClamp, aR, yClamp);
        }
    }

}

function _drawGridlinesMode(group, bounds, strokeColor, strokeWidth, gridDir, mode) {
    var m = (mode || 'straightLinesGrid').toLowerCase();
    
    if (m === 'withcircle') {
        _drawGridlinesFromSegments(group, strokeColor, strokeWidth, true, false, gridDir);
        return;
    }
    if (m === 'nocirclescomplex' || m === 'withoutcirclescomplex') {
        _drawGridlinesFromSegments(group, strokeColor, strokeWidth, false, true, gridDir);
        return;
    }
    if (m === 'circlegrid') {
        _drawCircleGrid(group, strokeColor, strokeWidth);
        return;
    }
    if (m === 'diagonalgrid') {
        _drawDiagonalGrid(group, strokeColor, strokeWidth);
        return;
    }
    if (m === 'straightlinesgrid') {
        _drawStraightLinesGrid(group, strokeColor, strokeWidth);
        return;
    }
    // default: simple
    _drawGridlines(group, bounds, strokeColor, strokeWidth, gridDir);
}

function _drawGridlinesFromSegments(group, strokeColor, strokeWidth, includeCircles, includeComplexTangents, gridDir) {
    var doc = app.activeDocument;
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
    var aL = ab[0], aT = ab[1], aR = ab[2], aB = ab[3];

    function addLine(x1, y1, x2, y2, key) {
        if (key && _seenLines[key]) return;
        if (key) _seenLines[key] = true;
        var ln = group.pathItems.add();
        ln.setEntirePath([[x1, y1], [x2, y2]]);
        ln.filled = false;
        ln.stroked = true;
        ln.strokeColor = strokeColor;
        ln.strokeWidth = strokeWidth;
        try { ln.name = 'Gridline'; } catch (e) {}
        return ln;
    }

    function addCircle(cx, cy, r, key) {
        if (key && _seenCircles[key]) return;
        if (key) _seenCircles[key] = true;
        var top = cy + r;
        var left = cx - r;
        var d = r * 2;
        var e = group.pathItems.ellipse(top, left, d, d);
        e.filled = false;
        e.stroked = true;
        e.strokeColor = strokeColor;
        e.strokeWidth = strokeWidth;
        try { e.name = 'Gridline'; } catch (e2) {}
        return e;
    }

    function round2(n) { return Math.round(n * 100) / 100; }
    function lineKey(x1,y1,x2,y2){
        // normalize order
        if (x1 > x2 || (x1 === x2 && y1 > y2)) {
            var tx=x1; x1=x2; x2=tx; var ty=y1; y1=y2; y2=ty;
        }
        return round2(x1)+','+round2(y1)+','+round2(x2)+','+round2(y2);
    }
    function circleKey(cx,cy,r){ return round2(cx)+','+round2(cy)+','+round2(r); }

    function vsub(a,b){ return [a[0]-b[0], a[1]-b[1]]; }
    function vlen(v){ return Math.sqrt(v[0]*v[0]+v[1]*v[1]); }
    function vnorm(v){ var l=vlen(v); return l<1e-6?[0,0]:[v[0]/l, v[1]/l]; }

    function almostEq(a,b,t){ return Math.abs(a-b) <= (t||0.01); }
    function ptEq(p,q,t){ return almostEq(p[0],q[0],t) && almostEq(p[1],q[1],t); }

    // intersection of infinite line with artboard rectangle -> return up to 2 pts
    function intersectLineWithRect(p0, p1) {
        var x0=p0[0], y0=p0[1], x1=p1[0], y1=p1[1];
        var pts=[];
        function add(x,y){
            // within rect
            if (x < aL-0.01 || x > aR+0.01 || y < aB-0.01 || y > aT+0.01) return;
            for (var i=0;i<pts.length;i++) if (ptEq(pts[i],[x,y],0.5)) return;
            pts.push([x,y]);
        }
        var dx=x1-x0, dy=y1-y0;
        if (Math.abs(dx) < 1e-9 && Math.abs(dy) < 1e-9) return pts;
        // x = aL
        if (Math.abs(dx) > 1e-9) {
            var t = (aL - x0)/dx; add(aL, y0 + t*dy);
            t = (aR - x0)/dx; add(aR, y0 + t*dy);
        }
        // y = aT / aB
        if (Math.abs(dy) > 1e-9) {
            var t2 = (aT - y0)/dy; add(x0 + t2*dx, aT);
            t2 = (aB - y0)/dy; add(x0 + t2*dx, aB);
        }
        return pts;
    }

    function bezierPoint(a0,h0,h1,a1,t){
        var mt=1-t;
        var mt2=mt*mt, t2=t*t;
        var x = a0[0]*mt2*mt + 3*h0[0]*mt2*t + 3*h1[0]*mt*t2 + a1[0]*t2*t;
        var y = a0[1]*mt2*mt + 3*h0[1]*mt2*t + 3*h1[1]*mt*t2 + a1[1]*t2*t;
        return [x,y];
    }

    function circleFrom3(p1,p2,p3){
        var x1=p1[0], y1=p1[1], x2=p2[0], y2=p2[1], x3=p3[0], y3=p3[1];
        var a = x1*(y2-y3) - y1*(x2-x3) + x2*y3 - x3*y2;
        if (Math.abs(a) < 1e-6) return null;
        var b = (x1*x1 + y1*y1)*(y3-y2) + (x2*x2 + y2*y2)*(y1-y3) + (x3*x3 + y3*y3)*(y2-y1);
        var c = (x1*x1 + y1*y1)*(x2-x3) + (x2*x2 + y2*y2)*(x3-x1) + (x3*x3 + y3*y3)*(x1-x2);
        var cx = -b/(2*a);
        var cy = -c/(2*a);
        var r = Math.sqrt((cx-x1)*(cx-x1) + (cy-y1)*(cy-y1));
        if (!isFinite(r) || r < 0.5) return null;
        return {cx:cx, cy:cy, r:r};
    }

    var dir = (gridDir || 'both').toLowerCase();
    var allowV = (dir === 'both' || dir === 'vertical' || dir === 'v');
    var allowH = (dir === 'both' || dir === 'horizontal' || dir === 'h');

    var _seenLines = {};
    var _seenCircles = {};

    if (!doc.selection || !doc.selection.length) return;

    for (var i=0;i<doc.selection.length;i++) {
        var paths = _collectPathItems(doc.selection[i]);
        for (var p=0;p<paths.length;p++) {
            var pi = paths[p];
            var pts = pi.pathPoints;
            var n = pts.length;
            for (var k=0;k<n;k++) {
                if (!pi.closed && k === n-1) break;
                var p0 = pts[k];
                var p1 = pts[(k+1)%n];

                var a0 = p0.anchor;
                var a1 = p1.anchor;
                var h0 = p0.rightDirection;
                var h1 = p1.leftDirection;

                var straight = ptEq(h0, a0, 0.05) && ptEq(h1, a1, 0.05);

                if (straight) {
                    // extend segment line to artboard bounds
                    var inter = intersectLineWithRect(a0, a1);
                    if (inter.length >= 2) {
                        // optional filter: if only V/H requested, skip other angles
                        var dx = a1[0]-a0[0];
                        var dy = a1[1]-a0[1];
                        if ((Math.abs(dx) < 1e-6 && allowV) || (Math.abs(dy) < 1e-6 && allowH) || (allowV && allowH)) {
                            var key = lineKey(inter[0][0], inter[0][1], inter[1][0], inter[1][1]);
                            addLine(inter[0][0], inter[0][1], inter[1][0], inter[1][1], key);
                        }
                    }
                } else {
                    // curve segment
                    if (includeComplexTangents) {
                        // tangent at start
                        var v0 = vsub(h0, a0);
                        if (vlen(v0) > 0.5) {
                            var t0 = [a0[0] + v0[0], a0[1] + v0[1]];
                            var inter0 = intersectLineWithRect(a0, t0);
                            if (inter0.length >= 2) {
                                var key0 = lineKey(inter0[0][0], inter0[0][1], inter0[1][0], inter0[1][1]);
                                addLine(inter0[0][0], inter0[0][1], inter0[1][0], inter0[1][1], key0);
                            }
                        }
                        // tangent at end
                        var v1 = vsub(a1, h1);
                        if (vlen(v1) > 0.5) {
                            var t1 = [a1[0] + v1[0], a1[1] + v1[1]];
                            var inter1 = intersectLineWithRect(a1, t1);
                            if (inter1.length >= 2) {
                                var key1 = lineKey(inter1[0][0], inter1[0][1], inter1[1][0], inter1[1][1]);
                                addLine(inter1[0][0], inter1[0][1], inter1[1][0], inter1[1][1], key1);
                            }
                        }
                    }

                    if (includeCircles) {
                        // approximate circle from 3 points on bezier
                        var q0 = bezierPoint(a0, h0, h1, a1, 0.0);
                        var q1 = bezierPoint(a0, h0, h1, a1, 0.5);
                        var q2 = bezierPoint(a0, h0, h1, a1, 1.0);
                        var c = circleFrom3(q0, q1, q2);
                        if (c && c.r < 1e6) {
                            var ck = circleKey(c.cx, c.cy, c.r);
                            addCircle(c.cx, c.cy, c.r, ck);
                        }
                    }
                }
            }
        }
    }
}

function _drawCircleGrid(group, strokeColor, strokeWidth) {
    // Reconstruct Circle Grid from Logo - detects circular Bezier segments and draws them.
    // No Artboard boundary restriction: circles are drawn at full radius and may extend beyond the Artboard (bleed out); not clipped or limited by Artboard dimensions.
    var doc = app.activeDocument;
    if (!doc.selection || !doc.selection.length) return;
    
    // Settings
    var tolerance = 1.2;       // max radial error in points (lower = stricter)
    var minRadius = 3;         // ignore tiny circles
    var dedupeCenterEps = 1.0; // points
    var dedupeRadiusEps = 1.0; // points
    
    function dist(a, b) {
        var dx = a[0] - b[0], dy = a[1] - b[1];
        return Math.sqrt(dx*dx + dy*dy);
    }
    
    // Cubic Bezier evaluation
    function bezierPoint(p0, p1, p2, p3, t) {
        var mt = 1 - t;
        var mt2 = mt * mt, t2 = t * t;
        var a = mt2 * mt;
        var b = 3 * mt2 * t;
        var c = 3 * mt * t2;
        var d = t2 * t;
        return [
            a*p0[0] + b*p1[0] + c*p2[0] + d*p3[0],
            a*p0[1] + b*p1[1] + c*p2[1] + d*p3[1]
        ];
    }
    
    // Circle from 3 points (circumcircle). Returns null if collinear.
    function circleFrom3(p1, p2, p3) {
        var x1=p1[0], y1=p1[1], x2=p2[0], y2=p2[1], x3=p3[0], y3=p3[1];
        var a = x1*(y2 - y3) - y1*(x2 - x3) + x2*y3 - x3*y2;
        if (Math.abs(a) < 1e-6) return null;
        
        var b = (x1*x1 + y1*y1)*(y3 - y2) + (x2*x2 + y2*y2)*(y1 - y3) + (x3*x3 + y3*y3)*(y2 - y1);
        var c = (x1*x1 + y1*y1)*(x2 - x3) + (x2*x2 + y2*y2)*(x3 - x1) + (x3*x3 + y3*y3)*(x1 - x2);
        
        var cx = -b / (2*a);
        var cy = -c / (2*a);
        var r = dist([cx, cy], p1);
        return { cx: cx, cy: cy, r: r };
    }
    
    function maxRadialError(circle, pts) {
        var maxErr = 0;
        for (var i=0; i<pts.length; i++) {
            var rr = dist([circle.cx, circle.cy], pts[i]);
            var err = Math.abs(rr - circle.r);
            if (err > maxErr) maxErr = err;
        }
        return maxErr;
    }
    
    function isStraight(p0, p1, p2, p3) {
        // If handles coincide with anchors, it's a straight segment in AI terms
        return (dist(p0, p1) < 0.001 && dist(p2, p3) < 0.001);
    }
    
    // Gather selected paths
    var paths = [];
    for (var s=0; s<doc.selection.length; s++) {
        var collected = _collectPathItems(doc.selection[s]);
        for (var ci=0; ci<collected.length; ci++) {
            paths.push(collected[ci]);
        }
    }
    if (paths.length === 0) return;
    
    // Find circle candidates from curved segments
    var circles = [];
    
    function addCircleCandidate(c) {
        if (!c || c.r < minRadius || !isFinite(c.r)) return;
        
        // Dedupe by near center/radius
        for (var i=0; i<circles.length; i++) {
            var o = circles[i];
            if (Math.abs(o.cx - c.cx) <= dedupeCenterEps &&
                Math.abs(o.cy - c.cy) <= dedupeCenterEps &&
                Math.abs(o.r - c.r) <= dedupeRadiusEps) {
                return;
            }
        }
        circles.push(c);
    }
    
    for (var p=0; p<paths.length; p++) {
        var it = paths[p];
        if (it.locked || it.hidden) continue;
        
        var pts = it.pathPoints;
        if (!pts || pts.length < 2) continue;
        
        var count = pts.length;
        var closed = it.closed;
        
        for (var i=0; i<count; i++) {
            var next = (i === count-1) ? 0 : i+1;
            if (!closed && i === count-1) break;
            
            var a0 = pts[i].anchor;
            var h1 = pts[i].rightDirection;
            var h2 = pts[next].leftDirection;
            var a3 = pts[next].anchor;
            
            if (isStraight(a0, h1, h2, a3)) continue;
            
            // Sample points on curve
            var p0 = a0;
            var pMid = bezierPoint(a0, h1, h2, a3, 0.5);
            var p3 = a3;
            
            var c = circleFrom3(p0, pMid, p3);
            if (!c) continue;
            
            // Validate with more samples
            var samples = [
                bezierPoint(a0, h1, h2, a3, 0.25),
                bezierPoint(a0, h1, h2, a3, 0.50),
                bezierPoint(a0, h1, h2, a3, 0.75)
            ];
            
            var err = maxRadialError(c, samples);
            if (err <= tolerance) addCircleCandidate(c);
        }
    }
    
    // Draw circles
    for (var cix=0; cix<circles.length; cix++) {
        var cc = circles[cix];
        var d = cc.r * 2;
        var top = cc.cy + cc.r;
        var left = cc.cx - cc.r;
        
        var circlePath = group.pathItems.ellipse(top, left, d, d);
        circlePath.stroked = true;
        circlePath.filled = false;
        circlePath.strokeWidth = strokeWidth;
        circlePath.strokeColor = strokeColor;
        circlePath.name = "GRID_CIRCLE_" + (cix + 1);
    }
}

function _lineArtboardIntersect(mx, my, dx, dy, aL, aT, aR, aB) {
    // Line: (mx, my) + t*(dx, dy). Return [[x1,y1],[x2,y2]] where line crosses artboard edges (no gaps at corners).
    var ts = [];
    if (Math.abs(dx) > 1e-9) {
        var tL = (aL - mx) / dx; var yL = my + tL * dy; if (yL >= aB - 0.01 && yL <= aT + 0.01) ts.push(tL);
        var tR = (aR - mx) / dx; var yR = my + tR * dy; if (yR >= aB - 0.01 && yR <= aT + 0.01) ts.push(tR);
    }
    if (Math.abs(dy) > 1e-9) {
        var tT = (aT - my) / dy; var xT = mx + tT * dx; if (xT >= aL - 0.01 && xT <= aR + 0.01) ts.push(tT);
        var tB = (aB - my) / dy; var xB = mx + tB * dx; if (xB >= aL - 0.01 && xB <= aR + 0.01) ts.push(tB);
    }
    if (ts.length < 2) return null;
    var tMin = Math.min.apply(null, ts);
    var tMax = Math.max.apply(null, ts);
    return [[mx + tMin * dx, my + tMin * dy], [mx + tMax * dx, my + tMax * dy]];
}

function _drawDiagonalGrid(group, strokeColor, strokeWidth) {
    // Diagonal Guides From Selected Logo - lines extend precisely to artboard edges (bounding box)
    var doc = app.activeDocument;
    if (!doc.selection || !doc.selection.length) return;
    
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    var aL = ab[0], aT = ab[1], aR = ab[2], aB = ab[3];
    
    var minSegLen = 6;          // ignore tiny segments (px)
    var ignoreNearDeg = 5;       // ignore near horizontal/vertical
    
    function dist(a, b, c, d) {
        var dx = c - a, dy = d - b;
        return Math.sqrt(dx*dx + dy*dy);
    }
    
    function angle(a, b, c, d) {
        var ang = Math.atan2(d - b, c - a) * 180 / Math.PI;
        if (ang < 0) ang += 180; // normalize to 0..180
        return ang;
    }
    
    function isNear(val, target, eps) {
        return Math.abs(val - target) <= eps;
    }
    
    function isStraightSegment(p0, p1) {
        var a0 = p0.anchor, r0 = p0.rightDirection;
        var l1 = p1.leftDirection, a1 = p1.anchor;
        return (dist(a0[0], a0[1], r0[0], r0[1]) < 0.01) && (dist(l1[0], l1[1], a1[0], a1[1]) < 0.01);
    }
    
    var paths = [];
    for (var s = 0; s < doc.selection.length; s++) {
        var collected = _collectPathItems(doc.selection[s]);
        for (var ci = 0; ci < collected.length; ci++) {
            paths.push(collected[ci]);
        }
    }
    if (paths.length === 0) return;
    
    for (var p = 0; p < paths.length; p++) {
        var path = paths[p];
        if (path.locked || path.hidden) continue;
        
        var pts = path.pathPoints;
        if (!pts || pts.length < 2) continue;
        
        var n = pts.length;
        
        function processPair(i, j) {
            var p0 = pts[i], p1 = pts[j];
            if (!isStraightSegment(p0, p1)) return;
            
            var a = p0.anchor;
            var b = p1.anchor;
            if (dist(a[0], a[1], b[0], b[1]) < minSegLen) return;
            
            var ang = angle(a[0], a[1], b[0], b[1]);
            if (isNear(ang, 0, ignoreNearDeg) || isNear(ang, 90, ignoreNearDeg) || isNear(ang, 180, ignoreNearDeg)) return;
            
            var mx = (a[0] + b[0]) / 2;
            var my = (a[1] + b[1]) / 2;
            var rad = ang * Math.PI / 180;
            var dx = Math.cos(rad), dy = Math.sin(rad);
            
            var seg = _lineArtboardIntersect(mx, my, dx, dy, aL, aT, aR, aB);
            if (!seg) return;
            
            var line = group.pathItems.add();
            line.setEntirePath([seg[0], seg[1]]);
            line.stroked = true;
            line.filled = false;
            line.strokeWidth = strokeWidth;
            line.strokeColor = strokeColor;
            line.name = "GRID_DIAGONAL_" + (p + 1) + "_" + (i + 1);
        }
        
        for (var i = 0; i < n - 1; i++) {
            processPair(i, i + 1);
        }
        if (path.closed && n > 2) {
            processPair(n - 1, 0);
        }
    }
}

function _drawStraightLinesGrid(group, strokeColor, strokeWidth) {
    // Straight (H/V) Guides From Selected Logo - detects straight horizontal and vertical segments
    var doc = app.activeDocument;
    if (!doc.selection || !doc.selection.length) return;
    
    // Settings
    var minSegLen = 4;            // ignore tiny segments
    var ignoreNearDeg = 4;        // angle tolerance for H/V detection
    var dedupeTol = 1.0;          // snap tolerance for unique X/Y lines (px)
    var extendFactor = 3.0;       // how far to extend beyond selection bounds
    
    function dist(a, b, c, d) {
        var dx = c - a, dy = d - b;
        return Math.sqrt(dx*dx + dy*dy);
    }
    
    function angleDeg(x1, y1, x2, y2) {
        var a = Math.atan2(y2 - y1, x2 - x1) * 180 / Math.PI;
        if (a < 0) a += 180; // 0..180
        return a;
    }
    
    function isNear(v, t, eps) {
        return Math.abs(v - t) <= eps;
    }
    
    function isStraightSegment(p0, p1) {
        // straight if handles sit on anchors
        var a0 = p0.anchor, r0 = p0.rightDirection;
        var l1 = p1.leftDirection, a1 = p1.anchor;
        return (dist(a0[0], a0[1], r0[0], r0[1]) < 0.01) && (dist(l1[0], l1[1], a1[0], a1[1]) < 0.01);
    }
    
    // Collect path items from selection
    var paths = [];
    for (var s = 0; s < doc.selection.length; s++) {
        var collected = _collectPathItems(doc.selection[s]);
        for (var ci = 0; ci < collected.length; ci++) {
            paths.push(collected[ci]);
        }
    }
    if (paths.length === 0) return;
    
    // Artboard bounds: vertical lines top→bottom, horizontal lines left→right, strictly within artboard
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
    var aL = ab[0], aT = ab[1], aR = ab[2], aB = ab[3];
    
    // Dedup helpers
    function snap(v) {
        return Math.round(v / dedupeTol) * dedupeTol;
    }
    var xs = {}; // snappedX -> true
    var ys = {}; // snappedY -> true
    
    // Extract H/V segments and store their mid X / mid Y
    for (var p = 0; p < paths.length; p++) {
        var path = paths[p];
        if (path.locked || path.hidden) continue;
        
        var pts = path.pathPoints;
        if (!pts || pts.length < 2) continue;
        
        var n = pts.length;
        
        function processPair(i, j) {
            var p0 = pts[i], p1 = pts[j];
            if (!isStraightSegment(p0, p1)) return;
            
            var a = p0.anchor;
            var b = p1.anchor;
            
            var len = dist(a[0], a[1], b[0], b[1]);
            if (len < minSegLen) return;
            
            var ang = angleDeg(a[0], a[1], b[0], b[1]);
            
            // Horizontal ~0/180
            if (isNear(ang, 0, ignoreNearDeg) || isNear(ang, 180, ignoreNearDeg)) {
                var yMid = (a[1] + b[1]) / 2;
                ys[snap(yMid)] = true;
                return;
            }
            
            // Vertical ~90
            if (isNear(ang, 90, ignoreNearDeg)) {
                var xMid = (a[0] + b[0]) / 2;
                xs[snap(xMid)] = true;
                return;
            }
        }
        
        // consecutive pairs
        for (var i = 0; i < n - 1; i++) {
            processPair(i, i + 1);
        }
        
        // closing segment
        if (path.closed && n > 2) {
            processPair(n - 1, 0);
        }
    }
    
    // Draw unique vertical lines: top of artboard → bottom of artboard, x clamped to artboard
    for (var xKey in xs) {
        var xVal = Number(xKey);
        var xClamp = Math.max(aL, Math.min(aR, xVal));
        var line = group.pathItems.add();
        line.setEntirePath([[xClamp, aT], [xClamp, aB]]);
        line.stroked = true;
        line.filled = false;
        line.strokeWidth = strokeWidth;
        line.strokeColor = strokeColor;
        line.name = "GRID_STRAIGHT_V_" + xKey;
    }
    
    // Draw unique horizontal lines: left of artboard → right of artboard, y clamped to artboard
    for (var yKey in ys) {
        var yVal = Number(yKey);
        var yClamp = Math.max(aB, Math.min(aT, yVal));
        var line = group.pathItems.add();
        line.setEntirePath([[aL, yClamp], [aR, yClamp]]);
        line.stroked = true;
        line.filled = false;
        line.strokeWidth = strokeWidth;
        line.strokeColor = strokeColor;
        line.name = "GRID_STRAIGHT_H_" + yKey;
    }
}

// =================== End Missing Helpers (v8.4.3 hotfix) ===================

// ---- Explicit exports for CEP (prevents "Error 24: ... is not a function") ----
// In some CEP/Illustrator setups, functions must be attached to $.global to be reliably callable.
// Functions are already attached above, but this ensures they're available even if loaded in different order.
(function() {
try {
        // Ensure $.global exists
        if (typeof $.global === 'undefined') {
            try {
                $.global = {};
            } catch (e) {
                // If we can't create $.global, functions won't be accessible from CEP
                return;
            }
        }
        
        // Explicitly attach all functions to $.global
        if (typeof helpersGenerateLogoGrid === 'function') {
            $.global.helpersGenerateLogoGrid = helpersGenerateLogoGrid;
        }
        if (typeof helpersGenerateClearspace === 'function') {
            $.global.helpersGenerateClearspace = helpersGenerateClearspace;
        }
        if (typeof helpersGenerateBaseGrid === 'function') {
            $.global.helpersGenerateBaseGrid = helpersGenerateBaseGrid;
        }
        if (typeof helpersMakeGuides === 'function') {
            $.global.helpersMakeGuides = helpersMakeGuides;
        }
        if (typeof helpersCleanup === 'function') {
            $.global.helpersCleanup = helpersCleanup;
        }
        if (typeof helpersCleanupByMode === 'function') {
            $.global.helpersCleanupByMode = helpersCleanupByMode;
        }
        if (typeof helpersPing === 'function') {
            $.global.helpersPing = helpersPing;
        }
        
        // Mark as loaded for bootstrap verification
        if (!$.global.Helpers) {
            $.global.Helpers = {};
    }
        $.global.Helpers.__loaded = true;
        $.global.__HELPERS_LOADED__ = true;
} catch (e) {
        // Log error but don't throw (allows panel to still open)
        try {
            $.writeln('Helpers: Error attaching functions to $.global: ' + e.toString());
        } catch (e2) {}
}
})();

// Color Sampling Function for Eyedropper
function helpersSampleColor() {
    try {
        if (!app.documents.length) {
            return JSON.stringify({ success: false, error: "No document open" });
        }
        
        var doc = app.activeDocument;
        var sampledColor = null;
        
        // Try to get color from selection (fill or stroke)
        if (doc.selection && doc.selection.length > 0) {
            var item = doc.selection[0];
            
            // Try fill color first
            if (item.filled && item.fillColor) {
                var fillColor = item.fillColor;
                if (fillColor.typename === "RGBColor") {
                    var r = Math.round(fillColor.red);
                    var g = Math.round(fillColor.green);
                    var b = Math.round(fillColor.blue);
                    sampledColor = "#" + 
                        ("0" + r.toString(16)).slice(-2) + 
                        ("0" + g.toString(16)).slice(-2) + 
                        ("0" + b.toString(16)).slice(-2);
                } else if (fillColor.typename === "CMYKColor") {
                    // Convert CMYK to RGB
                    var c = fillColor.cyan / 100;
                    var m = fillColor.magenta / 100;
                    var y = fillColor.yellow / 100;
                    var k = fillColor.black / 100;
                    var r = Math.round(255 * (1 - c) * (1 - k));
                    var g = Math.round(255 * (1 - m) * (1 - k));
                    var b = Math.round(255 * (1 - y) * (1 - k));
                    sampledColor = "#" + 
                        ("0" + r.toString(16)).slice(-2) + 
                        ("0" + g.toString(16)).slice(-2) + 
                        ("0" + b.toString(16)).slice(-2);
                } else if (fillColor.typename === "GrayColor") {
                    var gray = Math.round(fillColor.gray);
                    sampledColor = "#" + 
                        ("0" + gray.toString(16)).slice(-2) + 
                        ("0" + gray.toString(16)).slice(-2) + 
                        ("0" + gray.toString(16)).slice(-2);
                }
            }
            
            // If no fill, try stroke
            if (!sampledColor && item.stroked && item.strokeColor) {
                var strokeColor = item.strokeColor;
                if (strokeColor.typename === "RGBColor") {
                    var r = Math.round(strokeColor.red);
                    var g = Math.round(strokeColor.green);
                    var b = Math.round(strokeColor.blue);
                    sampledColor = "#" + 
                        ("0" + r.toString(16)).slice(-2) + 
                        ("0" + g.toString(16)).slice(-2) + 
                        ("0" + b.toString(16)).slice(-2);
                } else if (strokeColor.typename === "CMYKColor") {
                    // Convert CMYK to RGB
                    var c = strokeColor.cyan / 100;
                    var m = strokeColor.magenta / 100;
                    var y = strokeColor.yellow / 100;
                    var k = strokeColor.black / 100;
                    var r = Math.round(255 * (1 - c) * (1 - k));
                    var g = Math.round(255 * (1 - m) * (1 - k));
                    var b = Math.round(255 * (1 - y) * (1 - k));
                    sampledColor = "#" + 
                        ("0" + r.toString(16)).slice(-2) + 
                        ("0" + g.toString(16)).slice(-2) + 
                        ("0" + b.toString(16)).slice(-2);
                } else if (strokeColor.typename === "GrayColor") {
                    var gray = Math.round(strokeColor.gray);
                    sampledColor = "#" + 
                        ("0" + gray.toString(16)).slice(-2) + 
                        ("0" + gray.toString(16)).slice(-2) + 
                        ("0" + gray.toString(16)).slice(-2);
                }
            }
        }
        
        if (sampledColor) {
            return JSON.stringify({ success: true, color: sampledColor.toUpperCase() });
        } else {
            return JSON.stringify({ 
                success: false, 
                error: "Please select an object with a fill or stroke color, then click the eyedropper again." 
            });
        }
    } catch (e) {
        return JSON.stringify({ success: false, error: e.toString() });
    }
}
// Explicitly attach to $.global for CEP access
try { if (typeof $.global !== 'undefined') $.global.helpersSampleColor = helpersSampleColor; } catch(e) {}
