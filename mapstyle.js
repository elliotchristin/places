/* Map styles: OpenFreeMap "liberty" vector data, recoloured to feel like Apple Maps,
   plus a satellite (hybrid) style. Exposes window.MapStyles.build(mode) -> Promise<style>. */
'use strict';

(() => {
  const BASE_URL = 'https://tiles.openfreemap.org/styles/liberty';
  const SAT_TILES = 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}';

  const THEMES = {
    light: {
      land: '#f6f4ef', urban: '#efece6', park: '#cfe8c0', wood: '#c6e3b5', ice: '#ffffff', sand: '#f4ecd6',
      landuse: '#e9efe1', water: '#a7d3f5', building: '#e7e3dc', buildingLine: '#ddd8cf',
      road: '#ffffff', secondary: '#ffffff', primary: '#fff3c4', motorway: '#f8cf6b',
      casing: '#dcd6cc', motorwayCasing: '#e3b553', path: '#e2ddd3', rail: '#cfc9c0',
      boundary: '#b8aac9', label: '#3c3c43', halo: 'rgba(255,255,255,0.9)', waterLabel: '#3f7cb4', country: '#5a5a63',
      relief: 0.18,
    },
    dark: {
      land: '#212225', urban: '#26272b', park: '#1f3325', wood: '#1d3022', ice: '#3a3b3f', sand: '#2d2b26',
      landuse: '#242a24', water: '#1c2c40', building: '#2d2e32', buildingLine: '#35363b',
      road: '#3b3c41', secondary: '#45464c', primary: '#54555b', motorway: '#7a6438',
      casing: '#2a2b2f', motorwayCasing: '#5c4b2a', path: '#34353a', rail: '#3a3b40',
      boundary: '#6c6480', label: '#d6d6db', halo: 'rgba(20,20,22,0.85)', waterLabel: '#7fa6cf', country: '#bdbdc4',
      relief: 0,
    },
  };

  const HIDE = /^(park_outline|road_area_pattern|road_one_way_arrow.*|building-3d|poi_r7|poi_r20|highway-shield.*|road_shield_us|tunnel_.*_hatching)$/;
  const FRENCH = ['coalesce', ['get', 'name:fr'], ['get', 'name:latin'], ['get', 'name']];

  let basePromise = null;
  const loadBase = () => (basePromise ||= fetch(BASE_URL).then((r) => r.json()));
  const clone = (o) => JSON.parse(JSON.stringify(o));

  function lineColor(id, P) {
    if (/rail/.test(id)) return P.rail;
    if (/casing/.test(id)) return /motorway/.test(id) ? P.motorwayCasing : P.casing;
    if (/motorway/.test(id)) return P.motorway;
    if (/trunk_primary/.test(id)) return P.primary;
    if (/secondary_tertiary/.test(id)) return P.secondary;
    if (/path_pedestrian/.test(id)) return P.path;
    return P.road;
  }

  function recolor(style, P) {
    for (const l of style.layers) {
      const id = l.id;
      l.paint = l.paint || {};
      l.layout = l.layout || {};
      if (HIDE.test(id)) { l.layout.visibility = 'none'; continue; }
      if (l.type === 'background') l.paint['background-color'] = P.land;
      else if (id === 'natural_earth') { l.paint['raster-opacity'] = P.relief; if (!P.relief) l.layout.visibility = 'none'; }
      else if (l.type === 'fill') {
        const c = id === 'water' ? P.water
          : id === 'park' || id === 'landcover_grass' || id === 'landcover_wetland' ? P.park
          : id === 'landcover_wood' ? P.wood
          : id === 'landcover_ice' ? P.ice
          : id === 'landcover_sand' ? P.sand
          : id === 'landuse_residential' ? P.urban
          : id === 'building' ? P.building
          : id === 'aeroway_fill' ? P.landuse
          : /^landuse_/.test(id) ? P.landuse : null;
        if (c) { l.paint['fill-color'] = c; if (id !== 'water') l.paint['fill-opacity'] = 1; }
        if (id === 'building') l.paint['fill-outline-color'] = P.buildingLine;
      } else if (l.type === 'line') {
        if (/^waterway/.test(id)) l.paint['line-color'] = P.water;
        else if (/^boundary/.test(id)) l.paint['line-color'] = P.boundary;
        else if (/^aeroway/.test(id)) l.paint['line-color'] = P.road;
        else if (/^(road|bridge|tunnel)_/.test(id)) {
          l.paint['line-color'] = lineColor(id, P);
          if (/^tunnel_/.test(id)) l.paint['line-opacity'] = 0.5;
        }
      } else if (l.type === 'symbol') {
        if (l.layout['text-field'] && JSON.stringify(l.layout['text-field']).includes('name')) l.layout['text-field'] = FRENCH;
        l.paint['text-color'] = /^water|waterway/.test(id) ? P.waterLabel : /^label_country/.test(id) ? P.country : P.label;
        l.paint['text-halo-color'] = P.halo;
        l.paint['text-halo-width'] = 1.4;
      }
    }
    return style;
  }

  function satellite(base) {
    const s = clone(base);
    s.sources.satellite = {
      type: 'raster', tiles: [SAT_TILES], tileSize: 256, maxzoom: 19,
      attribution: 'Imagerie © Esri, Maxar, Earthstar Geographics',
    };
    const keep = /^(boundary_2|boundary_3|label_.*|water_name_.*|highway-name-major|airport)$/;
    s.layers = [
      { id: 'satellite', type: 'raster', source: 'satellite' },
      ...s.layers.filter((l) => keep.test(l.id)).map((l) => {
        l.paint = l.paint || {}; l.layout = l.layout || {};
        if (l.type === 'symbol') {
          if (l.layout['text-field']) l.layout['text-field'] = FRENCH;
          l.paint['text-color'] = '#ffffff';
          l.paint['text-halo-color'] = 'rgba(0,0,0,0.75)';
          l.paint['text-halo-width'] = 1.3;
        } else if (l.type === 'line') {
          l.paint['line-color'] = 'rgba(255,255,255,0.55)';
        }
        return l;
      }),
    ];
    return s;
  }

  window.MapStyles = {
    async build(mode) {
      const base = await loadBase();
      if (mode === 'satellite') return satellite(base);
      return recolor(clone(base), THEMES[mode] || THEMES.light);
    },
  };
})();
