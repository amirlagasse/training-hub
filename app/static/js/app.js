    const WORKOUT_TYPES = [
      ['Run', 'run'], ['Bike', 'bike'], ['Swim', 'swim'], ['Brick', 'brick'],
      ['Crosstrain', 'pulse'], ['Day Off', 'rest'], ['Mtn Bike', 'mtb'], ['Strength', 'strength'],
      ['Custom', 'timer'], ['XC-Ski', 'ski'], ['Rowing', 'rowing'], ['Walk', 'walk'],
      ['Other', 'other'],
    ];

    const OTHER_TYPES = [
      ['Event', 'event', 'event'],
      ['Goals', 'goal', 'goal'],
      ['Note', 'note', 'note'],
      ['Metrics', 'metrics', 'metrics'],
      ['Availability', 'calendar', 'availability'],
    ];

    const ICON_ASSETS = {
      run: '/icons/workouts/run.png',
      bike: '/icons/workouts/bike.png',
      swim: '/icons/workouts/swim.png',
      brick: '/icons/workouts/brick.png',
      pulse: '/icons/workouts/crosstrain.png',
      rest: '/icons/workouts/day_off.png',
      mtb: '/icons/workouts/mountian_bike.png',
      strength: '/icons/workouts/strength.png',
      timer: '/icons/workouts/other_custom.png',
      ski: '/icons/workouts/XC_Ski.png',
      rowing: '/icons/workouts/row.png',
      walk: '/icons/workouts/walk.png',
      other: '/icons/workouts/other_custom.png',
      event: '/icons/workouts/event.png',
      goal: '/icons/workouts/goal.png',
      note: '/icons/workouts/note.png',
      metrics: '/icons/workouts/note.png',
      calendar: '/icons/workouts/note.png',
    };

    const DOW = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];

    let activities = [];
    let calendarItems = [];
    let pairs = [];
    let selectedDate = todayKey();
    let selectedKind = 'workout';
    let selectedWorkoutType = 'Run';
    let editingItemId = null;
    let analyzeState = null;
    let initialMonthCentered = false;
    let appSettings = { units: { distance: 'km', elevation: 'm' }, ftp: {} };
    let distanceUnit = localStorage.getItem('distanceUnit') || 'km';
    let elevationUnit = localStorage.getItem('elevationUnit') || 'm';
    let currentFeel = 0;
    let fitUploadTargetActivityId = null;
    let fitUploadContext = 'global';
    let modalDraft = null;
    let calendarCursor = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
    let calendarScrollBound = false;
    let calendarAnchorWeekStart = dateKeyFromDate(mondayOfDate(new Date()));
    let calendarScrollTop = 0;

    function localDateKey(d) {
      const dt = new Date(d);
      const y = dt.getFullYear();
      const m = String(dt.getMonth() + 1).padStart(2, '0');
      const day = String(dt.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }

    function todayKey() {
      return localDateKey(new Date());
    }

    function parseDateKey(key) {
      return new Date(key + 'T00:00:00');
    }

    function dateKeyFromDate(d) {
      return localDateKey(d);
    }

    function mondayOfDate(dt) {
      const d = new Date(dt);
      d.setHours(0, 0, 0, 0);
      d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
      return d;
    }

    function syncCalendarHeaderFromScroll() {
      const wrap = document.getElementById('calendarScroll');
      if (!wrap) return;
      const rows = Array.from(wrap.querySelectorAll('.week-row'));
      if (!rows.length) return;
      const top = wrap.getBoundingClientRect().top;
      let closest = null;
      let closestAbove = null;
      rows.forEach((row) => {
        const r = row.getBoundingClientRect();
        if (r.top <= top + 2) closestAbove = row;
        if (!closest && r.bottom >= top) closest = row;
      });
      const row = closestAbove || closest || rows[0];
      if (!row) return;
      document.getElementById('calHeaderMonth').textContent = row.dataset.weekLabel || '';
      const [yy, mm] = String(row.dataset.weekMonth || '').split('-');
      if (yy && mm) calendarCursor = new Date(Number(yy), Number(mm) - 1, 1);
      calendarAnchorWeekStart = String(row.dataset.weekStart || '');
      calendarScrollTop = wrap.scrollTop;
    }

    function toDisplayDistanceFromMeters(meters, unit = distanceUnit) {
      const m = Number(meters || 0);
      if (unit === 'm') return { value: m, unit: 'm' };
      if (unit === 'mi') return { value: m / 1609.344, unit: 'mi' };
      return { value: m / 1000, unit: 'km' };
    }

    function toDisplayDistanceFromKm(km, unit = distanceUnit) {
      return toDisplayDistanceFromMeters(Number(km || 0) * 1000, unit);
    }

    function fromDisplayDistanceToKm(val, unit = distanceUnit) {
      const n = Number(val || 0);
      if (!Number.isFinite(n)) return 0;
      if (unit === 'm') return n / 1000;
      if (unit === 'mi') return n * 1.609344;
      return n;
    }

    function fromDisplayDistanceToMeters(val, unit = distanceUnit) {
      return fromDisplayDistanceToKm(val, unit) * 1000;
    }

    function toDisplayElevationFromMeters(meters, unit = elevationUnit) {
      const m = Number(meters || 0);
      if (unit === 'ft') return { value: m * 3.28084, unit: 'ft' };
      return { value: m, unit: 'm' };
    }

    function fromDisplayElevationToMeters(val, unit = elevationUnit) {
      const n = Number(val || 0);
      if (!Number.isFinite(n)) return 0;
      if (unit === 'ft') return n / 3.28084;
      return n;
    }

    function fmtDistanceMeters(meters) {
      const d = toDisplayDistanceFromMeters(meters, distanceUnit);
      return `${d.value.toFixed(distanceUnit === 'm' ? 0 : 1)} ${d.unit}`;
    }

    function fmtDistanceMetersInUnit(meters, unit) {
      const useUnit = unit || distanceUnit;
      const d = toDisplayDistanceFromMeters(meters, useUnit);
      return `${d.value.toFixed(useUnit === 'm' ? 0 : 1)} ${d.unit}`;
    }

    function fmtDistanceKm(km) {
      const d = toDisplayDistanceFromKm(km, distanceUnit);
      return `${d.value.toFixed(distanceUnit === 'm' ? 0 : 1)} ${d.unit}`;
    }

    function fmtHours(seconds) {
      return ((seconds || 0) / 3600).toFixed(1) + ' h';
    }

    function fmtDateLabel(key) {
      return parseDateKey(key).toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
    }

    function fmtDateUpper(key) {
      return parseDateKey(key).toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }).toUpperCase();
    }

    function fmtElevation(meters) {
      const e = toDisplayElevationFromMeters(meters, elevationUnit);
      return `${Math.round(e.value)} ${e.unit}`;
    }

    function monthKey(year, month) {
      return `${year}-${String(month + 1).padStart(2, '0')}`;
    }

    function isCalendarActive() {
      const node = document.getElementById('view-calendar');
      return !!(node && node.classList.contains('active'));
    }

    function setView(name) {
      document.querySelectorAll('.tab').forEach(el => el.classList.toggle('active', el.dataset.view === name));
      document.querySelectorAll('.view').forEach(el => el.classList.remove('active'));
      document.getElementById('view-' + name).classList.add('active');
      const pageHead = document.querySelector('.page-head');
      if (pageHead) pageHead.classList.toggle('hidden', name === 'calendar');
      document.getElementById('pageTitle').textContent = name.charAt(0).toUpperCase() + name.slice(1);
      if (name === 'calendar') {
        if (!calendarAnchorWeekStart) {
          jumpToCurrentMonth();
          return;
        }
        renderCalendar({ forceAnchor: false, preserveScroll: true });
      }
    }

    function updateUnitButtons() {
      document.querySelectorAll('.distance-unit-label').forEach((el) => { el.textContent = distanceUnit; });
      document.querySelectorAll('.elevation-unit-label').forEach((el) => { el.textContent = elevationUnit; });
    }

    function intensityByType(type) {
      const map = {
        Run: 0.85, Bike: 0.82, Swim: 0.8, Brick: 0.9, Crosstrain: 0.7, 'Day Off': 0.2,
        'Mtn Bike': 0.86, Strength: 0.75, Custom: 0.72, 'XC-Ski': 0.88, Rowing: 0.84,
        Walk: 0.55, Other: 0.65, Ride: 0.82, Workout: 0.8,
      };
      return map[type] || 0.7;
    }

    function activitySportKey(activity) {
      const t = String(activity.type || activity.sport_key || '').toLowerCase();
      if (t.includes('ride') || t.includes('cycle') || t.includes('bike')) return 'ride';
      if (t.includes('run') || t.includes('walk')) return 'run';
      if (t.includes('swim')) return 'swim';
      if (t.includes('row')) return 'row';
      if (t.includes('strength')) return 'strength';
      return 'other';
    }

    function ftpForActivity(activity) {
      const key = activitySportKey(activity);
      const ftp = Number((appSettings.ftp || {})[key] || 0);
      return ftp > 0 ? ftp : null;
    }

    function estimateTss(durationMin, intensity) {
      const durH = Math.max(0, Number(durationMin || 0)) / 60;
      const ifac = Math.max(0.2, Number(intensity || 0.7));
      return Math.round(durH * ifac * ifac * 100);
    }

    function activityToTss(activity) {
      if (Number(activity.tss_override || 0) > 0) return Number(activity.tss_override);
      const ifv = activityIF(activity);
      const durationH = Number(activity.moving_time || 0) / 3600;
      if (ifv && durationH > 0) return durationH * ifv * ifv * 100;
      return estimateTss(Number(activity.moving_time || 0) / 60, intensityByType(activity.type || 'Other'));
    }

    function itemToTss(item) {
      if (item.kind !== 'workout') return 0;
      const plannedTss = Number(item.planned_tss || 0);
      if (plannedTss > 0) return plannedTss;
      const userIntensity = Number(item.intensity || 0);
      const intensity = userIntensity > 0 ? (0.4 + Math.min(10, userIntensity) / 10) : intensityByType(item.workout_type || 'Other');
      return estimateTss(Number(item.duration_min || 0), intensity);
    }

    function plannedIF(item) {
      const ifv = Number(item.planned_if || 0);
      if (ifv > 0) return ifv;
      const tss = Number(item.planned_tss || 0);
      const hours = Number(item.duration_min || 0) / 60;
      if (tss > 0 && hours > 0) return Math.sqrt(tss / (hours * 100));
      return null;
    }

    function completedIF(obj) {
      const ifv = Number(obj.if_value || obj.completed_if || 0);
      if (ifv > 0) return ifv;
      const tss = Number(obj.tss_override || obj.completed_tss || 0);
      const hours = Number(obj.moving_time || (Number(obj.completed_duration_min || 0) * 60) || 0) / 3600;
      if (tss > 0 && hours > 0) return Math.sqrt(tss / (hours * 100));
      const ftp = ftpForActivity(obj);
      const avgP = Number(obj.avg_power || 0);
      if (ftp && avgP > 0) return avgP / ftp;
      return null;
    }

    function activityIF(activity) {
      return completedIF(activity);
    }

    function pairForPlanned(plannedId) {
      return pairs.find(p => p.planned_id === plannedId) || null;
    }

    function pairForStrava(stravaId) {
      return pairs.find(p => p.strava_id === String(stravaId)) || null;
    }

    function plannedMetric(plannedItem, basis) {
      if (basis === 'distance') return Number(plannedItem.distance_km || 0);
      if (basis === 'tss') return itemToTss(plannedItem);
      return Number(plannedItem.duration_min || 0);
    }

    function completedFromPlanned(plannedItem) {
      const dur = Number(plannedItem.completed_duration_min || 0);
      const dist = Number(plannedItem.completed_distance_km || 0);
      const tss = Number(plannedItem.completed_tss || 0);
      const ifv = Number(plannedItem.completed_if || 0);
      if (dur <= 0 && dist <= 0 && tss <= 0 && ifv <= 0) return null;
      return {
        moving_time: dur * 60,
        distance: dist * 1000,
        tss_override: tss,
        if_value: ifv,
        type: plannedItem.workout_type || 'Workout',
      };
    }

    function completedMetric(completedItem, basis) {
      if (basis === 'distance') return Number(completedItem.distance || 0) / 1000;
      if (basis === 'tss') {
        const override = Number(completedItem.tss_override || 0);
        if (override > 0) return override;
        const ifv = Number(completedItem.if_value || 0);
        const h = Number(completedItem.moving_time || 0) / 3600;
        if (ifv > 0 && h > 0) return h * ifv * ifv * 100;
        return activityToTss(completedItem);
      }
      return Number(completedItem.moving_time || 0) / 60;
    }

    function complianceStatus(plannedItem, completedItem, dayKey) {
      const today = todayKey();
      if (!plannedItem && completedItem) return { cls: 'unplanned', arrow: '' };
      if (plannedItem && !completedItem) {
        if (dayKey < today) return { cls: 'paired-red', arrow: '' };
        return { cls: 'workout', arrow: '' };
      }
      if (!plannedItem || !completedItem) return { cls: 'workout', arrow: '' };
      const hasPlannedBasis = ['duration', 'distance', 'tss'].some((basis) => plannedMetric(plannedItem, basis) > 0);
      if (!hasPlannedBasis) return { cls: 'unplanned', arrow: '' };
      const bases = ['duration', 'distance', 'tss'];
      const pcts = [];
      for (const basis of bases) {
        const p = plannedMetric(plannedItem, basis);
        const c = completedMetric(completedItem, basis);
        if (p > 0) {
          pcts.push((c / p) * 100);
        }
      }
      if (!pcts.length) return { cls: 'unplanned', arrow: '' };
      const best = pcts.reduce((bestPct, pct) => Math.abs(pct - 100) < Math.abs(bestPct - 100) ? pct : bestPct, pcts[0]);
      if (best >= 80 && best <= 120) return { cls: 'paired-green', arrow: '' };
      if ((best >= 50 && best < 80) || (best > 120 && best <= 150)) {
        return { cls: 'paired-yellow', arrow: best > 120 ? 'up' : 'down' };
      }
      return { cls: 'paired-orange', arrow: best > 120 ? 'up' : 'down' };
    }

    function buildObservedDailyTssMap() {
      const map = {};
      const today = todayKey();

      activities.forEach(a => {
        const key = dateKeyFromDate(new Date(a.start_date_local));
        map[key] = (map[key] || 0) + activityToTss(a);
      });

      return map;
    }

    function buildMetricsToDate(endKey) {
      const observed = buildObservedDailyTssMap();
      const endDate = parseDateKey(endKey);
      const today = parseDateKey(todayKey());
      const start = new Date(endDate);
      start.setDate(endDate.getDate() - 119);
      const values = [];

      for (let i = 0; i < 120; i += 1) {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        const key = dateKeyFromDate(d);
        values.push(d <= today ? Number(observed[key] || 0) : 0);
      }

      const ctlSeries = [];
      const atlSeries = [];
      const tsbSeries = [];
      let ctlPrev = 0;
      let atlPrev = 0;
      for (let i = 0; i < values.length; i += 1) {
        const tss = values[i];
        const tsb = ctlPrev - atlPrev;
        const ctl = ctlPrev + (tss - ctlPrev) / 42;
        const atl = atlPrev + (tss - atlPrev) / 7;
        tsbSeries.push(tsb);
        ctlSeries.push(ctl);
        atlSeries.push(atl);
        ctlPrev = ctl;
        atlPrev = atl;
      }

      return {
        ctl: Math.round(ctlSeries[ctlSeries.length - 1] || 0),
        atl: Math.round(atlSeries[atlSeries.length - 1] || 0),
        tsb: Math.round(tsbSeries[tsbSeries.length - 1] || 0),
        ctlSeries,
        atlSeries,
        tsbSeries,
      };
    }

    function renderSparkline(elId, series) {
      const el = document.getElementById(elId);
      el.innerHTML = '';
      if (!series.length) return;
      const recent = series.slice(-30);
      const maxVal = Math.max(...recent.map(v => Math.abs(v)), 1);
      recent.forEach(v => {
        const bar = document.createElement('span');
        bar.style.height = `${Math.max(3, Math.round((Math.abs(v) / maxVal) * 38))}px`;
        if (v < 0) bar.style.background = '#d5936f';
        el.appendChild(bar);
      });
    }

    function renderPerformanceMetrics() {
      const metrics = buildMetricsToDate(todayKey());
      document.getElementById('ctlVal').textContent = String(metrics.ctl);
      document.getElementById('atlVal').textContent = String(metrics.atl);
      document.getElementById('tsbVal').textContent = metrics.tsb > 0 ? `+${metrics.tsb}` : String(metrics.tsb);
      document.getElementById('ctlTrend').textContent = String(metrics.ctl);
      document.getElementById('atlTrend').textContent = String(metrics.atl);
      document.getElementById('tsbTrend').textContent = metrics.tsb > 0 ? `+${metrics.tsb}` : String(metrics.tsb);
      renderSparkline('ctlSpark', metrics.ctlSeries);
      renderSparkline('atlSpark', metrics.atlSeries);
      renderSparkline('tsbSpark', metrics.tsbSeries);
    }

    function iconSvg(name) {
      const src = ICON_ASSETS[name] || ICON_ASSETS.other;
      return `<span class="type-icon"><img src="${src}" alt="${name}"/></span>`;
    }

    function workoutIconKey(type) {
      const t = String(type || '').toLowerCase();
      if (t.includes('brick')) return 'brick';
      if (t.includes('cross')) return 'pulse';
      if (t.includes('day off') || t.includes('rest')) return 'rest';
      if (t.includes('mountain') || t.includes('mtn')) return 'mtb';
      if (t.includes('custom')) return 'timer';
      if (t.includes('ski')) return 'ski';
      if (t.includes('row')) return 'rowing';
      if (t.includes('run') || t.includes('walk')) return 'run';
      if (t.includes('swim')) return 'swim';
      if (t.includes('strength')) return 'strength';
      if (t.includes('ride') || t.includes('bike') || t.includes('cycling')) return 'bike';
      return 'other';
    }

    function cardIcon(type) {
      const key = workoutIconKey(type);
      const src = ICON_ASSETS[key] || ICON_ASSETS.other;
      return `<span class="wc-icon"><img src="${src}" alt="${key}" /></span>`;
    }

    function feelEmoji(v) {
      const map = { 1: '😫', 2: '🙁', 3: '😐', 4: '🙂', 5: '😁' };
      return map[Number(v)] || '';
    }

    function setFeelValue(v) {
      currentFeel = Number(v || 0);
      document.querySelectorAll('.feel-btn').forEach((btn) => {
        btn.classList.toggle('active', Number(btn.dataset.feel) === currentFeel);
      });
    }

    function commentsArrayFromEntity(entity) {
      if (!entity) return [];
      if (Array.isArray(entity.comments_feed)) {
        return entity.comments_feed.map((x) => String(x || '').trim()).filter(Boolean);
      }
      if (typeof entity.comments === 'string' && entity.comments.trim()) return [entity.comments.trim()];
      return [];
    }

    function renderCommentsFeed() {
      const wrap = document.getElementById('wvCommentsFeed');
      if (!modalDraft) {
        wrap.innerHTML = '';
        return;
      }
      const list = modalDraft.commentsFeed || [];
      if (!list.length) {
        wrap.innerHTML = '';
        return;
      }
      wrap.innerHTML = list.map((c, i) => `
        <div class="comment-item" data-index="${i}">
          <span class="comment-text" title="Double-click to edit">${c.replace(/</g, '&lt;')}</span>
          <button class="comment-delete" type="button" aria-label="Delete comment">🗑</button>
        </div>
      `).join('');
    }

    function commentCount(entity) {
      return commentsArrayFromEntity(entity).length;
    }

    function formatStartClock(isoText) {
      if (!isoText) return '--:--';
      const dt = new Date(isoText);
      if (Number.isNaN(dt.getTime())) return '--:--';
      return dt.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }).toLowerCase();
    }

    function parseDurationToMin(text) {
      const raw = String(text || '').trim();
      if (!raw) return 0;
      if (raw.includes(':')) {
        const parts = raw.split(':').map((x) => Number(x || 0));
        if (parts.length === 2) return parts[0] * 60 + parts[1];
        if (parts.length === 3) return parts[0] * 60 + parts[1] + (parts[2] / 60);
      }
      const n = Number(raw);
      return Number.isFinite(n) ? n : 0;
    }

    function formatDurationClock(mins) {
      const totalSec = Math.round(Math.max(0, Number(mins || 0)) * 60);
      const h = Math.floor(totalSec / 3600);
      const m = Math.floor((totalSec % 3600) / 60);
      const s = totalSec % 60;
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
    }

    function formatDurationClockCompact(mins) {
      return formatDurationClock(mins);
    }

    function toSportIcon(typeLabel) {
      const key = workoutIconKey(typeLabel);
      const src = ICON_ASSETS[key] || ICON_ASSETS.other;
      return `<span class="type-icon"><img src="${src}" alt="${key}" /></span>`;
    }

    function hasAnyCompletedMetric(obj, source) {
      if (source === 'strava') return true;
      return Number(obj.completed_duration_min || 0) > 0
        || Number(obj.completed_distance_m || 0) > 0
        || Number(obj.completed_distance_km || 0) > 0
        || Number(obj.completed_tss || 0) > 0
        || Number(obj.completed_if || 0) > 0;
    }

    function hasCompletedData(planned, completed) {
      if (completed) return true;
      if (!planned) return false;
      return Number(planned.completed_duration_min || 0) > 0
        || Number(planned.completed_distance_km || 0) > 0
        || Number(planned.completed_tss || 0) > 0;
    }

    function buildTypeGrids() {
      const workoutGrid = document.getElementById('workoutTypeGrid');
      workoutGrid.innerHTML = '';
      WORKOUT_TYPES.forEach(([name, icon]) => {
        const btn = document.createElement('button');
        btn.className = 'type-btn';
        btn.innerHTML = `${iconSvg(icon)}<span>${name}</span>`;
        btn.addEventListener('click', async () => {
          selectedKind = 'workout';
          selectedWorkoutType = name;
          const draft = {
            kind: 'workout',
            date: selectedDate || todayKey(),
            title: `Untitled ${name} Workout`,
            workout_type: name,
            duration_min: 0,
            distance_km: 0,
            intensity: 6,
          };
          closeActionModal();
          openWorkoutModal({ source: 'planned', data: draft, planned: draft, isDraft: true });
        });
        workoutGrid.appendChild(btn);
      });

      const otherGrid = document.getElementById('otherTypeGrid');
      otherGrid.innerHTML = '';
      OTHER_TYPES.forEach(([name, icon, kind]) => {
        const btn = document.createElement('button');
        btn.className = 'type-btn';
        btn.innerHTML = `${iconSvg(icon)}<span>${name}</span>`;
        btn.addEventListener('click', () => {
          selectedKind = kind;
          selectedWorkoutType = 'Other';
          openDetailModal();
        });
        otherGrid.appendChild(btn);
      });
    }

    function openActionModal(dateKey, forcedKind) {
      selectedDate = dateKey || todayKey();
      document.getElementById('actionDateTitle').textContent = fmtDateLabel(selectedDate);
      document.getElementById('actionDateTitle').classList.toggle('small', window.innerWidth < 1200);
      document.getElementById('actionModal').classList.add('open');
      if (forcedKind === 'event') {
        selectedKind = 'event';
        openDetailModal();
      }
      if (forcedKind === 'goal') {
        selectedKind = 'goal';
        openDetailModal();
      }
    }

    function closeActionModal() {
      document.getElementById('actionModal').classList.remove('open');
    }

    function openDetailModal(existingItem) {
      closeActionModal();
      const metrics = buildMetricsToDate(selectedDate);
      document.getElementById('detailDateLabel').textContent = fmtDateUpper(selectedDate);
      document.getElementById('miniCtl').textContent = `Fitness ${metrics.ctl}`;
      document.getElementById('miniAtl').textContent = `Fatigue ${metrics.atl}`;
      document.getElementById('miniTsb').textContent = `Form ${metrics.tsb > 0 ? '+' + metrics.tsb : metrics.tsb}`;

      const titleMap = {
        workout: 'Workout Title',
        event: 'Event Name',
        goal: 'Goal',
        note: 'Note Title',
        metrics: 'Metrics Entry',
        availability: 'Availability Title',
      };

      document.getElementById('detailTitleLabel').textContent = titleMap[selectedKind] || 'Title';
      editingItemId = existingItem ? existingItem.id : null;
      document.getElementById('deleteDetail').style.visibility = editingItemId ? 'visible' : 'hidden';

      document.getElementById('dDate').value = existingItem ? existingItem.date : selectedDate;
      document.getElementById('dTitle').value = existingItem ? (existingItem.title || '') : '';
      const detailDesc = existingItem ? (existingItem.description || '') : '';
      document.getElementById('dDescription').value = detailDesc;
      document.getElementById('dDescriptionOther').value = detailDesc;
      document.getElementById('dWorkoutType').value = existingItem ? (existingItem.workout_type || selectedWorkoutType) : selectedWorkoutType;
      document.getElementById('dDuration').value = existingItem ? (existingItem.duration_min || '') : '';
      if (existingItem && Number(existingItem.distance_km || 0) > 0) {
        document.getElementById('dDistance').value = toDisplayDistanceFromKm(existingItem.distance_km).value.toFixed(distanceUnit === 'm' ? 0 : 1);
      } else {
        document.getElementById('dDistance').value = '';
      }
      document.getElementById('dIntensity').value = existingItem ? (existingItem.intensity || 6) : '6';
      document.getElementById('dEventType').value = existingItem ? (existingItem.event_type || 'Race') : 'Race';
      document.getElementById('dAvailability').value = existingItem ? (existingItem.availability || 'Unavailable') : 'Unavailable';

      document.getElementById('workoutFields').classList.toggle('hidden', selectedKind !== 'workout');
      document.getElementById('eventFields').classList.toggle('hidden', selectedKind !== 'event');
      document.getElementById('availabilityFields').classList.toggle('hidden', selectedKind !== 'availability');
      const isWorkout = selectedKind === 'workout';
      document.getElementById('detailMetricsChips').style.display = isWorkout ? 'contents' : 'none';
      document.querySelector('.detail-right').style.display = isWorkout ? 'block' : 'none';
      document.querySelector('.detail-body').style.gridTemplateColumns = isWorkout ? '1fr 340px' : '1fr';
      document.getElementById('nonWorkoutDescription').style.display = isWorkout ? 'none' : 'block';

      document.getElementById('detailModal').classList.add('open');
    }

    function closeDetailModal() {
      document.getElementById('detailModal').classList.remove('open');
    }

    async function saveDetail(closeAfter) {
      const payload = {
        kind: selectedKind,
        date: document.getElementById('dDate').value,
        title: document.getElementById('dTitle').value,
        description: selectedKind === 'workout'
          ? document.getElementById('dDescription').value
          : document.getElementById('dDescriptionOther').value,
      };

      if (selectedKind === 'workout') {
        payload.workout_type = document.getElementById('dWorkoutType').value || selectedWorkoutType;
        payload.duration_min = Number(document.getElementById('dDuration').value || 0);
        payload.distance_km = fromDisplayDistanceToKm(document.getElementById('dDistance').value);
        payload.intensity = Number(document.getElementById('dIntensity').value || 6);
      }

      if (selectedKind === 'event') {
        payload.event_type = document.getElementById('dEventType').value;
      }

      if (selectedKind === 'availability') {
        payload.availability = document.getElementById('dAvailability').value;
      }

      const url = editingItemId ? `/calendar-items/${editingItemId}` : '/calendar-items';
      const method = editingItemId ? 'PUT' : 'POST';
      const resp = await fetch(url, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      if (!resp.ok) {
        const err = await resp.text();
        console.error('Could not save item:', err);
        return;
      }

      await loadData(false);
      if (closeAfter) {
        closeDetailModal();
      }
    }

    async function deleteCurrentDetail() {
      if (!editingItemId) return;
      if (!window.confirm('Are you sure you want to delete this?')) return;
      const resp = await fetch(`/calendar-items/${editingItemId}`, { method: 'DELETE' });
      if (!resp.ok) return;
      closeDetailModal();
      await loadData(false);
    }

    function buildDayAggregateMap() {
      const map = {};
      const pairPlanned = new Set(pairs.map(p => String(p.planned_id)));

      activities.forEach(a => {
        const key = dateKeyFromDate(new Date(a.start_date_local));
        if (!map[key]) {
          map[key] = { done: [], items: [], durationMin: 0, tss: 0 };
        }
        map[key].done.push(a);
        map[key].durationMin += Number(a.moving_time || 0) / 60;
        map[key].tss += activityToTss(a);
      });

      calendarItems.forEach(item => {
        const key = item.date;
        if (!map[key]) {
          map[key] = { done: [], items: [], durationMin: 0, tss: 0 };
        }
        map[key].items.push(item);
        if (item.kind === 'workout' && !pairPlanned.has(String(item.id))) {
          const manual = completedFromPlanned(item);
          if (manual) {
            map[key].durationMin += Number(manual.moving_time || 0) / 60;
            map[key].tss += Number(manual.tss_override || 0) || activityToTss(manual);
          }
        }
      });

      return map;
    }

    function formatDurationMin(mins) {
      const total = Math.max(0, Math.round(mins));
      const h = Math.floor(total / 60);
      const m = total % 60;
      return `${h}:${String(m).padStart(2, '0')}`;
    }

    function getWeekMetrics(dateKeys, dayMap) {
      let durationMin = 0;
      let tss = 0;
      let weekEnd = null;

      dateKeys.forEach(key => {
        if (!key) return;
        weekEnd = key;
        const day = dayMap[key];
        if (!day) return;
        durationMin += day.durationMin;
        tss += day.tss;
      });

      const metrics = weekEnd ? buildMetricsToDate(weekEnd) : { ctl: 0, atl: 0, tsb: 0 };
      return {
        durationLabel: formatDurationMin(durationMin),
        tss: Math.round(tss),
        ctl: metrics.ctl,
        atl: metrics.atl,
        tsb: metrics.tsb,
      };
    }

    function closeContextMenu() {
      const menu = document.getElementById('contextMenu');
      menu.style.display = 'none';
      menu.innerHTML = '';
      menu.dataset.itemId = '';
    }

    function openContextMenu(x, y, options) {
      const menu = document.getElementById('contextMenu');
      menu.innerHTML = '';
      options.forEach(opt => {
        const btn = document.createElement('button');
        btn.textContent = opt.label;
        btn.addEventListener('click', async () => {
          closeContextMenu();
          await opt.onClick();
        });
        menu.appendChild(btn);
      });
      menu.style.left = `${x}px`;
      menu.style.top = `${y}px`;
      menu.style.display = 'block';
    }

    function showItemMenu(ev, payload) {
      ev.preventDefault();
      ev.stopPropagation();
      const opts = [];
      if (payload.source === 'planned') {
        opts.push({
          label: 'Edit',
          onClick: async () => {
            if (payload.data.kind === 'workout') {
              openWorkoutModal({ source: 'planned', data: payload.data, planned: payload.data });
            } else {
              selectedKind = payload.data.kind || 'workout';
              selectedDate = payload.data.date;
              selectedWorkoutType = payload.data.workout_type || 'Other';
              openDetailModal(payload.data);
            }
          },
        });
        opts.push({
          label: 'Copy',
          onClick: async () => {
            const copy = { ...payload.data };
            delete copy.id;
            delete copy.created_at;
            copy.title = `${copy.title || 'Copy'} (Copy)`;
            await fetch('/calendar-items', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(copy),
            });
            await loadData(false);
          },
        });
        opts.push({
          label: 'Delete',
          onClick: async () => {
            if (!window.confirm('Are you sure you want to delete this?')) return;
            await fetch(`/calendar-items/${payload.data.id}`, { method: 'DELETE' });
            await loadData(false);
          },
        });
      }
      if (payload.source === 'strava') {
        opts.push({
          label: 'Delete',
          onClick: async () => {
            if (!window.confirm('Are you sure you want to delete this?')) return;
            await fetch(`/activities/${payload.data.id}`, { method: 'DELETE' });
            await loadData(false);
          },
        });
      }

      const plannedId = payload.source === 'planned' ? payload.data.id : null;
      const stravaId = payload.source === 'strava' ? String(payload.data.id) : null;
      const currentPair = plannedId ? pairForPlanned(plannedId) : stravaId ? pairForStrava(stravaId) : null;
      if (currentPair) {
        opts.push({
          label: 'Unpair',
          onClick: async () => {
            await fetch(`/pairs/${currentPair.id}`, { method: 'DELETE' });
            await loadData(false);
          },
        });
      }

      if (!opts.length) return;
      openContextMenu(ev.clientX, ev.clientY, opts);
    }

    async function pairWorkouts(plannedId, stravaId) {
      if (!plannedId || !stravaId) return;
      const planned = calendarItems.find(i => String(i.id) === String(plannedId));
      const completed = activities.find(a => String(a.id) === String(stravaId));
      const typeLabel = (completed && completed.type) ? completed.type : (planned && planned.workout_type) ? planned.workout_type : 'Workout';
      const untitled = `Untitled ${typeLabel} Workout`;
      await fetch('/pairs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          planned_id: plannedId,
          strava_id: String(stravaId),
          override_date: planned ? planned.date : '',
          override_title: untitled,
        }),
      });
      await loadData(false);
    }

    function updateRpeLabel() {
      const slider = document.getElementById('wvRpe');
      const label = document.getElementById('wvRpeVal');
      if (!modalDraft || slider.disabled || !modalDraft.rpeTouched) {
        label.textContent = '';
        return;
      }
      label.textContent = String(slider.value || '--');
    }

    function setWorkoutMode(mode) {
      const analyzeBtn = document.getElementById('wvAnalyzeBtn');
      const modal = document.getElementById('workoutViewModal');
      const showAnalyze = mode === 'analyze' && !analyzeBtn.classList.contains('disabled');
      document.getElementById('wvSummary').classList.toggle('hidden', showAnalyze);
      document.getElementById('wvAnalyze').classList.toggle('hidden', !showAnalyze);
      analyzeBtn.classList.toggle('active', showAnalyze);
      analyzeBtn.textContent = showAnalyze ? 'Summary' : 'Analyze';
      analyzeBtn.classList.toggle('summary', showAnalyze);
      analyzeBtn.classList.toggle('analyze', !showAnalyze);
      if (modal) modal.classList.toggle('analyze-mode', showAnalyze);
      document.getElementById('wvFilesPopover').classList.add('hidden');
      document.getElementById('wvFilesTabBtn').classList.remove('active');
    }

    function syncUnitSelectValue(selectId, desired, allowed) {
      const el = document.getElementById(selectId);
      if (!el) return;
      if (allowed.includes(desired)) {
        el.value = desired;
      } else {
        el.value = allowed[0];
      }
    }

    function recalcIfTssRows() {
      const parseDurHours = (id) => parseDurationToMin(document.getElementById(id).value) / 60;
      const recalcPair = (durId, tssId, ifId) => {
        const hours = parseDurHours(durId);
        const tss = Number(document.getElementById(tssId).value || 0);
        const ifv = Number(document.getElementById(ifId).value || 0);
        if (hours <= 0) return;
        if (ifv > 0) {
          document.getElementById(tssId).value = (hours * ifv * ifv * 100).toFixed(1);
        } else if (tss > 0) {
          document.getElementById(ifId).value = Math.sqrt(tss / (hours * 100)).toFixed(2);
        }
      };
      recalcPair('pcDurPlan', 'pcTssPlan', 'pcIfPlan');
      recalcPair('pcDurComp', 'pcTssComp', 'pcIfComp');
    }

    function renderWorkoutFiles(payload) {
      const node = document.getElementById('wvFileList');
      const data = payload.data || {};
      const browseBtn = document.getElementById('wvBrowseFilesBtn');
      browseBtn.disabled = false;
      const rows = [];
      const hasExisting = !!data.fit_id;
      if (hasExisting && modalDraft && modalDraft.pendingDeleteFit) {
        rows.push(`
          <div class="wv-file-row">
            <div><strong>${data.fit_filename || `${data.fit_id}.fit`}</strong><div class="meta">Will be deleted on Save</div></div>
            <div class="wv-file-actions">
              <button class="btn secondary" id="wvUndoDeleteFitBtn">Undo</button>
            </div>
          </div>
        `);
      } else if (hasExisting) {
        const fileName = data.fit_filename || `${data.fit_id}.fit`;
        rows.push(`
          <div class="wv-file-row">
            <div><strong>${fileName}</strong><div class="meta">FIT attached</div></div>
            <div class="wv-file-actions">
              <button class="btn secondary" id="wvRecalcFitBtn">Recalculate</button>
              <button class="btn secondary" id="wvDeleteFitBtn">Delete</button>
              <a class="btn secondary" id="wvDownloadFitBtn" href="/activities/${encodeURIComponent(data.id)}/fit/download">Download</a>
            </div>
          </div>
        `);
      }
      if (!rows.length) {
        node.innerHTML = '<p class="meta">No files attached.</p>';
        return;
      }
      node.innerHTML = rows.join('');
      const undoBtn = document.getElementById('wvUndoDeleteFitBtn');
      if (undoBtn) {
        undoBtn.onclick = () => {
          if (!modalDraft) return;
          modalDraft.pendingDeleteFit = false;
          renderWorkoutFiles(payload);
        };
      }
      const recalcBtn = document.getElementById('wvRecalcFitBtn');
      if (recalcBtn) {
        recalcBtn.onclick = async () => {
          if (modalDraft && modalDraft.pendingDeleteFit) return;
          const resp = await fetch(`/activities/${data.id}/fit/recalculate`, { method: 'POST' });
          if (!resp.ok) return;
          const refreshed = await resp.json();
          window.currentWorkoutPayload = { ...payload, data: refreshed };
          renderWorkoutSummary(window.currentWorkoutPayload);
          await renderWorkoutAnalyze(window.currentWorkoutPayload);
          renderWorkoutFiles(window.currentWorkoutPayload);
          await loadData(false);
        };
      }
      const delBtn = document.getElementById('wvDeleteFitBtn');
      if (delBtn) {
        delBtn.onclick = () => {
          if (!modalDraft) return;
          if (!window.confirm('Are you sure you want to delete this?')) return;
          modalDraft.pendingDeleteFit = true;
          renderWorkoutFiles(window.currentWorkoutPayload);
        };
      }
    }

    function openWorkoutModal(payload) {
      window.currentWorkoutPayload = payload;
      const data = payload.data || {};
      const parentPlanned = payload.planned || null;
      modalDraft = {
        isNewWorkout: !!payload.isDraft,
        pendingDeleteFit: false,
        uploadedNow: false,
        createdActivityId: null,
        createdPairId: null,
        originalFit: payload.source === 'strava' ? {
          fit_id: data.fit_id || null,
          fit_filename: data.fit_filename || null,
          distance: data.distance || 0,
          moving_time: data.moving_time || 0,
          avg_power: data.avg_power || null,
          avg_hr: data.avg_hr || null,
          min_hr: data.min_hr || null,
          max_hr: data.max_hr || null,
          min_power: data.min_power || null,
          max_power: data.max_power || null,
          elev_gain_m: data.elev_gain_m || null,
          if_value: data.if_value || null,
          tss_override: data.tss_override || null,
        } : null,
        commentsFeed: commentsArrayFromEntity(parentPlanned || data),
        sportType: parentPlanned
          ? (parentPlanned.workout_type || 'Other')
          : payload.source === 'strava' ? (data.type || 'Other') : (data.workout_type || 'Other'),
        rpeTouched: false,
      };
      const typeLabel = parentPlanned
        ? (parentPlanned.workout_type || 'Workout')
        : payload.source === 'strava' ? (data.type || 'Workout') : (data.workout_type || 'Workout');
      const dateKey = parentPlanned ? parentPlanned.date : payload.source === 'strava' ? dateKeyFromDate(new Date(data.start_date_local)) : data.date;
      const metrics = buildMetricsToDate(dateKey || todayKey());
      const dateText = parseDateKey(dateKey || todayKey()).toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }).toUpperCase();
      document.getElementById('wvDateLine').textContent = dateText;
      document.getElementById('wvHeadCtl').textContent = `Fitness ${metrics.ctl}`;
      document.getElementById('wvHeadAtl').textContent = `Fatigue ${metrics.atl}`;
      document.getElementById('wvHeadTsb').textContent = `Form ${metrics.tsb > 0 ? '+' + metrics.tsb : metrics.tsb}`;
      document.getElementById('wvTimeSelect').innerHTML = `<option>${payload.source === 'strava' ? formatStartClock(data.start_date_local) : '8:00 am'}</option>`;
      document.getElementById('wvTitle').value = parentPlanned ? (parentPlanned.title || 'Workout') : (data.title || data.name || 'Workout');
      const subNode = document.getElementById('wvSub');
      if (subNode) subNode.textContent = `${typeLabel} • ${dateLabel}`;
      modalDraft.sportType = WORKOUT_TYPES.some(([n]) => n === modalDraft.sportType) ? modalDraft.sportType : 'Other';
      document.getElementById('wvSportIcon').innerHTML = toSportIcon(modalDraft.sportType);
      document.getElementById('wvSportName').textContent = modalDraft.sportType;
      document.getElementById('wvCommentInput').value = '';
      document.getElementById('deleteWorkoutView').style.visibility = (parentPlanned && parentPlanned.id) || data.id ? 'visible' : 'hidden';

      const analyzeToggle = document.getElementById('wvAnalyzeBtn');
      const hasFile = !!(data.fit_id);
      analyzeToggle.classList.toggle('disabled', !hasFile);
      const sportMenu = document.getElementById('wvSportMenu');
      sportMenu.innerHTML = WORKOUT_TYPES.map(([name, icon]) => `
        <button class="wv-sport-item" type="button" data-sport="${name}">
          ${iconSvg(icon)}
          <span>${name}</span>
        </button>
      `).join('');
      sportMenu.querySelectorAll('.wv-sport-item').forEach((btn) => {
        btn.addEventListener('click', () => {
          if (!modalDraft) return;
          modalDraft.sportType = btn.dataset.sport || 'Other';
          document.getElementById('wvSportIcon').innerHTML = toSportIcon(modalDraft.sportType);
          document.getElementById('wvSportName').textContent = modalDraft.sportType;
          sportMenu.classList.add('hidden');
        });
      });
      if (hasFile) {
        renderWorkoutAnalyze(payload);
      } else {
        document.getElementById('wvAnalyze').classList.add('hidden');
        document.getElementById('wvSelectionKv').innerHTML = '<div>No FIT stream for this workout.</div>';
      }
      renderWorkoutSummary(payload);
      renderWorkoutFiles(payload);
      setWorkoutMode('summary');
      document.getElementById('wvFilesPopover').classList.add('hidden');
      document.getElementById('wvFilesTabBtn').classList.remove('active');
      document.getElementById('wvSportMenu').classList.add('hidden');
      document.getElementById('workoutViewModal').classList.add('open');
    }

    async function closeWorkoutModal(discard = true) {
      const payload = window.currentWorkoutPayload;
      if (discard && modalDraft && payload) {
        const data = payload.data || {};
        try {
          if (modalDraft.uploadedNow) {
            if (modalDraft.createdPairId) {
              await fetch(`/pairs/${modalDraft.createdPairId}`, { method: 'DELETE' });
            }
            if (modalDraft.createdActivityId) {
              await fetch(`/activities/${modalDraft.createdActivityId}`, { method: 'DELETE' });
            } else if (payload.source === 'strava') {
              const oldFit = modalDraft.originalFit || {};
              if (oldFit.fit_id) {
                await fetch(`/activities/${data.id}/fit/restore`, {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify(oldFit),
                });
              } else if (data.fit_id) {
                await fetch(`/activities/${data.id}/fit`, { method: 'DELETE' });
              }
            }
          }
        } catch (_err) {
          // best-effort rollback
        }
      }
      document.getElementById('workoutViewModal').classList.remove('open');
      document.getElementById('workoutViewModal').classList.remove('analyze-mode');
      modalDraft = null;
      window.currentWorkoutPayload = null;
      fitUploadContext = 'global';
      fitUploadTargetActivityId = null;
    }

    async function saveWorkoutView(closeAfter) {
      const payload = window.currentWorkoutPayload;
      if (!payload) return;
      const data = payload.data || {};
      const targetPlanned = payload.planned || (payload.source === 'planned' ? data : null);
      recalcIfTssRows();
      const description = document.getElementById('wvDescription').value;
      const commentsFeed = modalDraft ? modalDraft.commentsFeed.slice() : [];
      const comments = commentsFeed.length ? commentsFeed[commentsFeed.length - 1] : '';
      const sport = (modalDraft && modalDraft.sportType) || 'Other';
      const distanceUnitLocal = document.getElementById('pcDistanceUnit').value || 'km';
      const elevationUnitLocal = document.getElementById('pcElevationUnit').value || 'm';
      const plannedDuration = parseDurationToMin(document.getElementById('pcDurPlan').value);
      const plannedDistanceM = fromDisplayDistanceToMeters(document.getElementById('pcDistPlan').value, distanceUnitLocal);
      const plannedElevationM = fromDisplayElevationToMeters(document.getElementById('pcElevPlan').value, elevationUnitLocal);
      const plannedTss = Number(document.getElementById('pcTssPlan').value || 0);
      const plannedIf = Number(document.getElementById('pcIfPlan').value || 0);
      const completedDuration = parseDurationToMin(document.getElementById('pcDurComp').value);
      const completedDistanceM = fromDisplayDistanceToMeters(document.getElementById('pcDistComp').value, distanceUnitLocal);
      const completedElevationM = fromDisplayElevationToMeters(document.getElementById('pcElevComp').value, elevationUnitLocal);
      const completedTss = Number(document.getElementById('pcTssComp').value || 0);
      const completedIf = Number(document.getElementById('pcIfComp').value || 0);
      const hasCompleted = completedDuration > 0 || completedDistanceM > 0 || completedTss > 0 || payload.source === 'strava';
      const hasAnyNumericValue = plannedDuration > 0
        || plannedDistanceM > 0
        || plannedElevationM > 0
        || plannedTss > 0
        || plannedIf > 0
        || completedDuration > 0
        || completedDistanceM > 0
        || completedElevationM > 0
        || completedTss > 0
        || completedIf > 0;
      const rpeVal = Number(document.getElementById('wvRpe').value || 0);
      const feel = hasCompleted ? currentFeel : 0;
      const rpeOut = (hasCompleted && modalDraft && modalDraft.rpeTouched) ? rpeVal : 0;
      const isNewDraft = !!(modalDraft && modalDraft.isNewWorkout);

      if (isNewDraft && !hasAnyNumericValue) {
        await closeWorkoutModal(false);
        return;
      }

      let activityData = payload.source === 'strava' ? data : null;
      if (activityData && modalDraft && modalDraft.pendingDeleteFit && activityData.fit_id) {
        const delResp = await fetch(`/activities/${activityData.id}/fit`, { method: 'DELETE' });
        if (delResp.ok) activityData = await delResp.json();
      }

      if (targetPlanned && targetPlanned.id) {
        await fetch(`/calendar-items/${targetPlanned.id}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ...targetPlanned,
            title: document.getElementById('wvTitle').value.trim() || targetPlanned.title || 'Untitled Workout',
            workout_type: sport,
            duration_min: plannedDuration,
            distance_km: plannedDistanceM / 1000,
            distance_m: plannedDistanceM,
            elevation_m: plannedElevationM,
            distance_unit: distanceUnitLocal,
            elevation_unit: elevationUnitLocal,
            planned_tss: plannedTss,
            planned_if: plannedIf,
            description,
            comments,
            comments_feed: commentsFeed,
            feel,
            rpe: rpeOut,
            completed_duration_min: completedDuration,
            completed_distance_km: completedDistanceM / 1000,
            completed_distance_m: completedDistanceM,
            completed_elevation_m: completedElevationM,
            completed_tss: completedTss,
            completed_if: completedIf,
          }),
        });
      } else if (payload.source === 'planned') {
        const createResp = await fetch('/calendar-items', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            kind: 'workout',
            date: data.date || selectedDate || todayKey(),
            title: document.getElementById('wvTitle').value.trim() || `Untitled ${sport} Workout`,
            workout_type: sport,
            duration_min: plannedDuration,
            distance_km: plannedDistanceM / 1000,
            distance_m: plannedDistanceM,
            elevation_m: plannedElevationM,
            distance_unit: distanceUnitLocal,
            elevation_unit: elevationUnitLocal,
            planned_tss: plannedTss,
            planned_if: plannedIf,
            description,
            comments,
            comments_feed: commentsFeed,
            feel,
            rpe: rpeOut,
            completed_duration_min: completedDuration,
            completed_distance_km: completedDistanceM / 1000,
            completed_distance_m: completedDistanceM,
            completed_elevation_m: completedElevationM,
            completed_tss: completedTss,
            completed_if: completedIf,
            intensity: Number(data.intensity || 6) || 6,
          }),
        });
        if (createResp.ok) {
          const created = await createResp.json();
          if (modalDraft) modalDraft.isNewWorkout = false;
          window.currentWorkoutPayload = { source: 'planned', data: created, planned: created };
        }
      } else if (payload.source === 'strava') {
        await fetch(`/activities/${data.id}/meta`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            description,
            comments,
            comments_feed: commentsFeed,
            feel,
            rpe: rpeOut,
            if_value: completedIf,
            tss_override: completedTss,
            title: document.getElementById('wvTitle').value.trim(),
            type: sport,
          }),
        });
      }
      await loadData(false);
      if (closeAfter || isNewDraft) await closeWorkoutModal(false);
    }

    function num(v) {
      const n = Number(v);
      return Number.isFinite(n) ? n : null;
    }

    function timeToSec(iso, baseMs) {
      const t = new Date(iso).getTime();
      return Math.max(0, (t - baseMs) / 1000);
    }

    function hms(totalSec) {
      const s = Math.max(0, Math.round(totalSec));
      const h = Math.floor(s / 3600);
      const m = Math.floor((s % 3600) / 60);
      const sec = s % 60;
      return `${h}:${String(m).padStart(2, '0')}:${String(sec).padStart(2, '0')}`;
    }

    function fmtAxis(val, key) {
      if (key === 'speed') {
        if (distanceUnit === 'mi') return `${(val * 2.23694).toFixed(1)} mph`;
        if (distanceUnit === 'm') return `${val.toFixed(2)} m/s`;
        return `${(val * 3.6).toFixed(1)} km/h`;
      }
      if (key === 'distance') return fmtDistanceMeters(val);
      if (key === 'power') return `${Math.round(val)} W`;
      if (key === 'heart_rate') return `${Math.round(val)} bpm`;
      if (key === 'cadence') return `${Math.round(val)} rpm`;
      if (key === 'altitude') return fmtElevation(val);
      return String(Math.round(val));
    }

    function renderWorkoutSummary(payload) {
      const data = payload.data || {};
      const parentPlanned = payload.planned || null;
      const explicitCompleted = parentPlanned ? completedFromPlanned(parentPlanned) : completedFromPlanned(data);
      const completedDurationMin = payload.source === 'strava'
        ? Number(data.moving_time || 0) / 60
        : explicitCompleted ? Number(explicitCompleted.moving_time || 0) / 60 : 0;
      const completedDistanceM = payload.source === 'strava'
        ? Number(data.distance || 0)
        : explicitCompleted ? Number(explicitCompleted.distance || 0) : Number((parentPlanned || data).completed_distance_m || 0);
      const completedTss = payload.source === 'strava'
        ? activityToTss(data)
        : explicitCompleted ? Number(explicitCompleted.tss_override || 0) : 0;
      const typeLabel = parentPlanned
        ? (parentPlanned.workout_type || 'Workout')
        : payload.source === 'strava' ? (data.type || 'Workout') : (data.workout_type || 'Workout');
      const dateLabel = parentPlanned
        ? `${parentPlanned.date} (Planned Day)`
        : payload.source === 'strava'
          ? new Date(data.start_date_local).toLocaleString()
          : `${data.date} (Planned)`;
      const plannedObj = parentPlanned || data;
      let distanceUnitLocal = String(plannedObj.distance_unit || distanceUnit || 'km');
      let elevationUnitLocal = String(plannedObj.elevation_unit || elevationUnit || 'm');
      syncUnitSelectValue('pcDistanceUnit', distanceUnitLocal, ['km', 'mi', 'm']);
      syncUnitSelectValue('pcElevationUnit', elevationUnitLocal, ['m', 'ft']);

      document.getElementById('wvSummaryText').textContent = '';
      document.getElementById('wvDescription').value = (parentPlanned && parentPlanned.description) || data.description || '';
      if (modalDraft) {
        if (!Array.isArray(modalDraft.commentsFeed) || !modalDraft.commentsFeed.length) {
          modalDraft.commentsFeed = commentsArrayFromEntity(parentPlanned || data);
        }
        renderCommentsFeed();
      }
      const savedRpe = Number((parentPlanned && parentPlanned.rpe) || data.rpe || 0);
      if (modalDraft) modalDraft.rpeTouched = savedRpe > 0;
      document.getElementById('wvRpe').value = String(savedRpe > 0 ? savedRpe : 1);
      document.getElementById('wvRpe').classList.toggle('rpe-unset', !(modalDraft && modalDraft.rpeTouched));
      setFeelValue((parentPlanned && parentPlanned.feel) || data.feel || 0);

      const plannedDuration = parentPlanned ? Number(parentPlanned.duration_min || 0) : Number(data.duration_min || 0);
      const plannedDistanceM = Number((plannedObj.distance_m || 0) || (Number(plannedObj.distance_km || 0) * 1000));
      const plannedElevationM = Number(plannedObj.elevation_m || 0);
      const completedElevationM = payload.source === 'strava'
        ? Number(data.elev_gain_m || 0)
        : Number(plannedObj.completed_elevation_m || 0);
      const plannedTss = Number(plannedObj.planned_tss || 0) || (parentPlanned ? itemToTss(parentPlanned) : itemToTss(data));
      const plannedIf = plannedIF(plannedObj);
      const completedIf = payload.source === 'strava' ? activityIF(data) : completedIF({
        completed_if: plannedObj.completed_if,
        completed_tss: completedTss,
        moving_time: completedDurationMin * 60,
        avg_power: data.avg_power,
        type: plannedObj.workout_type || data.type,
      });

      document.getElementById('pcDurPlan').value = plannedDuration ? formatDurationClock(plannedDuration) : '';
      document.getElementById('pcDurComp').value = completedDurationMin ? formatDurationClock(completedDurationMin) : '';
      const pd = toDisplayDistanceFromMeters(plannedDistanceM, distanceUnitLocal);
      const cd = toDisplayDistanceFromMeters(completedDistanceM, distanceUnitLocal);
      const pe = toDisplayElevationFromMeters(plannedElevationM, elevationUnitLocal);
      const ce = toDisplayElevationFromMeters(completedElevationM, elevationUnitLocal);
      document.getElementById('pcDistPlan').value = plannedDistanceM > 0 ? pd.value.toFixed(distanceUnitLocal === 'm' ? 0 : 1) : '';
      document.getElementById('pcDistComp').value = completedDistanceM > 0 ? cd.value.toFixed(distanceUnitLocal === 'm' ? 0 : 1) : '';
      document.getElementById('pcElevPlan').value = plannedElevationM > 0 ? pe.value.toFixed(elevationUnitLocal === 'm' ? 0 : 1) : '';
      document.getElementById('pcElevComp').value = completedElevationM > 0 ? ce.value.toFixed(elevationUnitLocal === 'm' ? 0 : 1) : '';
      document.getElementById('pcTssPlan').value = plannedTss ? String(Math.round(plannedTss)) : '';
      document.getElementById('pcTssComp').value = completedTss ? String(Math.round(completedTss)) : '';
      document.getElementById('pcIfPlan').value = plannedIf ? Number(plannedIf).toFixed(2) : '';
      document.getElementById('pcIfComp').value = completedIf ? Number(completedIf).toFixed(2) : '';
      document.getElementById('wvHrMin').value = data.min_hr ? String(Math.round(data.min_hr)) : '';
      document.getElementById('wvHrAvg').value = data.avg_hr ? String(Math.round(data.avg_hr)) : '';
      document.getElementById('wvHrMax').value = data.max_hr ? String(Math.round(data.max_hr)) : '';
      document.getElementById('wvPowerMin').value = data.min_power ? String(Math.round(data.min_power)) : '';
      document.getElementById('wvPowerAvg').value = data.avg_power ? String(Math.round(data.avg_power)) : '';
      document.getElementById('wvPowerMax').value = data.max_power ? String(Math.round(data.max_power)) : '';
      if (data.fit_id) {
        fetch(`/fit/${data.fit_id}`).then(r => r.ok ? r.json() : null).then((fit) => {
          if (!fit) return;
          const s = fit.summary || {};
          if (s.min_hr) document.getElementById('wvHrMin').value = String(Math.round(s.min_hr));
          if (s.avg_hr) document.getElementById('wvHrAvg').value = String(Math.round(s.avg_hr));
          if (s.max_hr) document.getElementById('wvHrMax').value = String(Math.round(s.max_hr));
          if (s.min_power) document.getElementById('wvPowerMin').value = String(Math.round(s.min_power));
          if (s.avg_power) document.getElementById('wvPowerAvg').value = String(Math.round(s.avg_power));
          if (s.max_power) document.getElementById('wvPowerMax').value = String(Math.round(s.max_power));
        }).catch(() => {});
      }

      const hasCompleted = hasAnyCompletedMetric((parentPlanned || data), payload.source);
      document.querySelectorAll('.feel-btn').forEach((btn) => { btn.disabled = !hasCompleted; });
      document.getElementById('wvRpe').disabled = !hasCompleted;
      if (!hasCompleted && modalDraft) modalDraft.rpeTouched = false;
      document.getElementById('wvRpe').classList.toggle('rpe-unset', !hasCompleted || !(modalDraft && modalDraft.rpeTouched));
      updateRpeLabel();

      document.getElementById('wvHeaderDuration').textContent = completedDurationMin ? formatDurationClock(completedDurationMin) : '--:--:--';
      document.getElementById('wvHeaderDistance').textContent = completedDistanceM ? fmtDistanceMeters(completedDistanceM) : '--';
      document.getElementById('wvHeaderTss').textContent = completedTss ? `${Math.round(completedTss)} TSS` : '-- TSS';

      ['pcDurComp', 'pcDistComp', 'pcElevComp', 'pcTssComp', 'pcIfComp', 'pcDurPlan', 'pcDistPlan', 'pcElevPlan', 'pcTssPlan', 'pcIfPlan'].forEach((id) => {
        const el = document.getElementById(id);
        el.classList.remove('muted');
        el.oninput = null;
      });

      const syncDistanceUnits = (nextUnit) => {
        const planMeters = fromDisplayDistanceToMeters(document.getElementById('pcDistPlan').value, distanceUnitLocal);
        const compMeters = fromDisplayDistanceToMeters(document.getElementById('pcDistComp').value, distanceUnitLocal);
        distanceUnitLocal = nextUnit;
        const nextPlan = toDisplayDistanceFromMeters(planMeters, nextUnit);
        const nextComp = toDisplayDistanceFromMeters(compMeters, nextUnit);
        document.getElementById('pcDistPlan').value = planMeters > 0 ? nextPlan.value.toFixed(nextUnit === 'm' ? 0 : 1) : '';
        document.getElementById('pcDistComp').value = compMeters > 0 ? nextComp.value.toFixed(nextUnit === 'm' ? 0 : 1) : '';
      };
      const syncElevationUnits = (nextUnit) => {
        const planMeters = fromDisplayElevationToMeters(document.getElementById('pcElevPlan').value, elevationUnitLocal);
        const compMeters = fromDisplayElevationToMeters(document.getElementById('pcElevComp').value, elevationUnitLocal);
        elevationUnitLocal = nextUnit;
        const nextPlan = toDisplayElevationFromMeters(planMeters, nextUnit);
        const nextComp = toDisplayElevationFromMeters(compMeters, nextUnit);
        document.getElementById('pcElevPlan').value = planMeters > 0 ? nextPlan.value.toFixed(nextUnit === 'm' ? 0 : 1) : '';
        document.getElementById('pcElevComp').value = compMeters > 0 ? nextComp.value.toFixed(nextUnit === 'm' ? 0 : 1) : '';
      };
      document.getElementById('pcDistanceUnit').onchange = (ev) => syncDistanceUnits(ev.target.value);
      document.getElementById('pcElevationUnit').onchange = (ev) => syncElevationUnits(ev.target.value);
      ['pcDurPlan', 'pcDurComp', 'pcTssPlan', 'pcTssComp', 'pcIfPlan', 'pcIfComp'].forEach((id) => {
        document.getElementById(id).onchange = recalcIfTssRows;
      });
    }

    async function renderWorkoutAnalyze(payload) {
      const data = payload.data || {};
      if (!data.fit_id) {
        document.getElementById('wvSelectionKv').innerHTML = '<div>No FIT stream for this workout.</div>';
        document.querySelector('#wvLapTable tbody').innerHTML = '';
        document.getElementById('wvChart').innerHTML = '';
        document.getElementById('wvViewfinder').innerHTML = '';
        return;
      }

      const resp = await fetch(`/fit/${data.fit_id}`);
      if (!resp.ok) {
        document.getElementById('wvSelectionKv').innerHTML = '<div>Could not load FIT data.</div>';
        return;
      }
      const fit = await resp.json();
      const series = Array.isArray(fit.series) ? fit.series : [];
      const laps = Array.isArray(fit.laps) ? fit.laps : [];
      const summary = fit.summary || {};
      if (!series.length) {
        document.getElementById('wvSelectionKv').innerHTML = '<div>No FIT points available.</div>';
        return;
      }

      const baseMs = new Date(series[0].timestamp).getTime();
      const pts = series.map((p) => ({
        t: timeToSec(p.timestamp, baseMs),
        heart_rate: num(p.heart_rate),
        speed: num(p.speed),
        distance: num(p.distance),
        cadence: num(p.cadence),
        power: num(p.power),
        altitude: num(p.altitude),
      }));
      const totalSec = Math.max(1, pts[pts.length - 1].t - pts[0].t);

      analyzeState = {
        pts,
        laps,
        totalSec,
        wStart: 0,
        wEnd: totalSec,
      };

      const chart = document.getElementById('wvChart');
      const finder = document.getElementById('wvViewfinder');
      const left = 54;
      const right = 170;
      const top = 14;
      const bottom = 28;
      const w = 1200;
      const h = 360;
      const cw = w - left - right;
      const ch = h - top - bottom;
      const lineMeta = [
        { key: 'heart_rate', color: '#f35353', label: 'HR' },
        { key: 'power', color: '#ff62f2', label: 'W' },
        { key: 'cadence', color: '#f39b1f', label: 'RPM' },
        { key: 'speed', color: '#3fa144', label: 'MPH' },
      ];

      function valPath(meta, inWindow) {
        const vals = inWindow.map(p => p[meta.key]).filter(v => v !== null);
        if (!vals.length) return { path: '', min: 0, max: 1, avg: null };
        const min = Math.min(...vals);
        const max = Math.max(...vals);
        const span = Math.max(0.001, max - min);
        const avg = vals.reduce((s, v) => s + v, 0) / vals.length;
        let started = false;
        let d = '';
        inWindow.forEach((p) => {
          const v = p[meta.key];
          if (v === null) return;
          const x = left + ((p.t - analyzeState.wStart) / Math.max(1, analyzeState.wEnd - analyzeState.wStart)) * cw;
          const y = top + (1 - ((v - min) / span)) * ch;
          d += `${started ? 'L' : 'M'}${x.toFixed(2)} ${y.toFixed(2)} `;
          started = true;
        });
        return { path: d.trim(), min, max, avg };
      }

      function inWindow() {
        return pts.filter(p => p.t >= analyzeState.wStart && p.t <= analyzeState.wEnd);
      }

      function renderSelectionStats() {
        const win = inWindow();
        const duration = Math.max(1, analyzeState.wEnd - analyzeState.wStart);
        const frac = duration / totalSec;
        const distance = Number(summary.distance_m || 0) * frac;
        const mean = (k) => {
          const vals = win.map(p => p[k]).filter(v => v !== null);
          return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
        };
        document.getElementById('wvSelectionKv').innerHTML = `
          <div>Duration<strong>${hms(duration)}</strong></div>
          <div>Distance<strong>${fmtDistanceMeters(distance)}</strong></div>
          <div>Avg HR<strong>${mean('heart_rate') ? Math.round(mean('heart_rate')) : '--'}</strong></div>
          <div>Avg Power<strong>${mean('power') ? `${Math.round(mean('power'))} W` : '--'}</strong></div>
          <div>Avg Cadence<strong>${mean('cadence') ? `${Math.round(mean('cadence'))} rpm` : '--'}</strong></div>
          <div>Avg Speed<strong>${mean('speed') ? fmtAxis(mean('speed'), 'speed') : '--'}</strong></div>
          <div>Elevation Gain<strong>${summary.elev_gain_m ? fmtElevation(summary.elev_gain_m) : '--'}</strong></div>
        `;
      }

      function renderMain() {
        const win = inWindow();
        const xTicks = 5;
        const paths = lineMeta.map(m => ({ ...m, ...valPath(m, win) }));
        let svg = `<rect x="0" y="0" width="${w}" height="${h}" fill="#f3f7fd" stroke="#d6e1ee"></rect>`;
        for (let i = 0; i <= xTicks; i += 1) {
          const x = left + (i / xTicks) * cw;
          svg += `<line x1="${x}" y1="${top}" x2="${x}" y2="${top + ch}" stroke="#e1eaf5"/>`;
          const sec = analyzeState.wStart + (i / xTicks) * (analyzeState.wEnd - analyzeState.wStart);
          svg += `<text x="${x}" y="${h - 8}" fill="#5b7290" font-size="11" text-anchor="middle">${hms(sec)}</text>`;
        }
        paths.forEach((p) => {
          if (p.path) svg += `<path d="${p.path}" stroke="${p.color}" stroke-width="2" fill="none"></path>`;
        });
        paths.forEach((p, idx) => {
          const y = top + 14 + idx * 24;
          svg += `<text x="${w - right + 6}" y="${y}" fill="${p.color}" font-size="11">${p.label} ${fmtAxis(p.max, p.key)} / ${fmtAxis(p.min, p.key)}</text>`;
        });
        chart.innerHTML = svg;
        renderSelectionStats();
      }

      function renderFinder() {
        const fw = 1200;
        const fh = 90;
        const px = (t) => (t / totalSec) * fw;
        const speedVals = pts.map(p => p.speed).filter(v => v !== null);
        const sMin = speedVals.length ? Math.min(...speedVals) : 0;
        const sMax = speedVals.length ? Math.max(...speedVals) : 1;
        const sSpan = Math.max(0.001, sMax - sMin);
        let d = '';
        let started = false;
        pts.forEach((p) => {
          if (p.speed === null) return;
          const x = px(p.t);
          const y = 6 + (1 - ((p.speed - sMin) / sSpan)) * (fh - 24);
          d += `${started ? 'L' : 'M'}${x.toFixed(2)} ${y.toFixed(2)} `;
          started = true;
        });
        const bx = px(analyzeState.wStart);
        const bw = Math.max(8, px(analyzeState.wEnd) - bx);
        finder.innerHTML = `
          <rect x="0" y="0" width="${fw}" height="${fh}" fill="#edf3fb" stroke="#d6e1ee"></rect>
          <path d="${d}" stroke="#3fa144" stroke-width="1.5" fill="none"></path>
          <rect id="wvBrush" x="${bx}" y="2" width="${bw}" height="${fh - 4}" fill="rgba(80,150,255,.22)" stroke="#2a66d2"></rect>
        `;
      }

      let dragging = false;
      let dragOffset = 0;
      finder.onmousedown = (ev) => {
        const rect = finder.getBoundingClientRect();
        const x = ((ev.clientX - rect.left) / rect.width) * 1200;
        const b = finder.querySelector('#wvBrush');
        const bx = Number(b.getAttribute('x'));
        const bw = Number(b.getAttribute('width'));
        if (x >= bx && x <= bx + bw) {
          dragging = true;
          dragOffset = x - bx;
        } else {
          const center = x / 1200;
          const span = (analyzeState.wEnd - analyzeState.wStart) / totalSec;
          let s = Math.max(0, center - span / 2);
          let e = Math.min(1, center + span / 2);
          if (e - s < span) s = Math.max(0, e - span);
          analyzeState.wStart = s * totalSec;
          analyzeState.wEnd = e * totalSec;
          renderFinder();
          renderMain();
        }
      };
      finder.onmousemove = (ev) => {
        if (!dragging) return;
        const rect = finder.getBoundingClientRect();
        const x = ((ev.clientX - rect.left) / rect.width) * 1200;
        const b = finder.querySelector('#wvBrush');
        const bw = Number(b.getAttribute('width'));
        let bx = x - dragOffset;
        bx = Math.max(0, Math.min(1200 - bw, bx));
        b.setAttribute('x', String(bx));
        analyzeState.wStart = (bx / 1200) * totalSec;
        analyzeState.wEnd = ((bx + bw) / 1200) * totalSec;
        renderMain();
      };
      window.onmouseup = () => { dragging = false; };

      const lapBody = document.querySelector('#wvLapTable tbody');
      lapBody.innerHTML = '';
      const lapRows = laps.length ? laps : [{
        name: 'Lap 1',
        start: series[0].timestamp,
        end: series[series.length - 1].timestamp,
        duration_s: totalSec,
      }];
      lapRows.forEach((lap, idx) => {
        const startSec = Math.max(0, timeToSec(lap.start || series[0].timestamp, baseMs));
        const endSec = Math.max(startSec, lap.end ? timeToSec(lap.end, baseMs) : (startSec + Number(lap.duration_s || 0)));
        const row = document.createElement('tr');
        row.innerHTML = `
          <td><input type="checkbox" /></td>
          <td>${lap.name || `Lap #${idx + 1}`}</td>
          <td>${hms(Number(lap.duration_s || (endSec - startSec)))}</td>
          <td>${lap.distance_m ? fmtDistanceMeters(lap.distance_m) : '--'}</td>
          <td>${lap.avg_hr ? Math.round(lap.avg_hr) : '--'}</td>
          <td>${lap.avg_power ? Math.round(lap.avg_power) : '--'}</td>
        `;
        row.querySelector('input').addEventListener('change', (ev) => {
          row.classList.toggle('selected', ev.target.checked);
          const checked = Array.from(lapBody.querySelectorAll('input')).map((el, i) => ({ checked: el.checked, i })).filter(x => x.checked);
          if (!checked.length) {
            analyzeState.wStart = 0;
            analyzeState.wEnd = totalSec;
          } else {
            const minI = Math.min(...checked.map(c => c.i));
            const maxI = Math.max(...checked.map(c => c.i));
            const minLap = lapRows[minI];
            const maxLap = lapRows[maxI];
            analyzeState.wStart = Math.max(0, timeToSec(minLap.start || series[0].timestamp, baseMs));
            analyzeState.wEnd = Math.max(analyzeState.wStart + 1, timeToSec(maxLap.end || series[series.length - 1].timestamp, baseMs));
          }
          renderFinder();
          renderMain();
        });
        lapBody.appendChild(row);
      });

      document.getElementById('wvHrMin').value = summary.min_hr ? String(Math.round(summary.min_hr)) : '';
      document.getElementById('wvHrAvg').value = summary.avg_hr ? String(Math.round(summary.avg_hr)) : '';
      document.getElementById('wvHrMax').value = summary.max_hr ? String(Math.round(summary.max_hr)) : '';
      document.getElementById('wvPowerMin').value = summary.min_power ? String(Math.round(summary.min_power)) : '';
      document.getElementById('wvPowerAvg').value = summary.avg_power ? String(Math.round(summary.avg_power)) : '';
      document.getElementById('wvPowerMax').value = summary.max_power ? String(Math.round(summary.max_power)) : '';
      renderFinder();
      renderMain();
    }

    function renderEvents() {
      const list = document.getElementById('eventsList');
      const events = calendarItems
        .filter(i => i.kind === 'event')
        .sort((a, b) => (a.date > b.date ? 1 : -1))
        .slice(0, 5);

      list.innerHTML = '';
      if (!events.length) {
        list.innerHTML = '<p class="meta">No events yet. Click + to add one.</p>';
        return;
      }

      events.forEach(e => {
        const node = document.createElement('div');
        node.className = 'event-item';
        node.innerHTML = `<h4>${e.title}</h4><p>${e.date} • ${e.event_type || 'Event'}</p>`;
        list.appendChild(node);
      });
    }

    function renderGoals() {
      const list = document.getElementById('goalsList');
      const goals = calendarItems
        .filter(i => i.kind === 'goal')
        .sort((a, b) => (a.date > b.date ? 1 : -1))
        .slice(0, 6);

      list.innerHTML = '';
      if (!goals.length) {
        list.innerHTML = '<p class="meta">No goals yet. Click Add Goal.</p>';
        return;
      }

      goals.forEach(g => {
        const node = document.createElement('div');
        node.className = 'goal-item';
        node.innerHTML = `<h4>${g.title}</h4><p>${g.date}</p>`;
        list.appendChild(node);
      });
    }

    function renderHome() {
      const today = todayKey();
      const doneToday = activities.filter(a => dateKeyFromDate(new Date(a.start_date_local)) === today);

      const plannedUpcoming = calendarItems
        .filter(i => i.kind === 'workout' && i.date >= today)
        .sort((a, b) => (a.date > b.date ? 1 : -1))
        .slice(0, 8);

      const doneNode = document.getElementById('todayDone');
      doneNode.innerHTML = '';
      if (!doneToday.length) {
        doneNode.innerHTML = '<p class="meta">No completed workouts for today yet.</p>';
      } else {
        doneToday.forEach(a => {
          const row = document.createElement('button');
          row.type = 'button';
          row.className = 'list-item';
          row.innerHTML = `
            <div>
              <p class="title">${a.name || 'Workout'}</p>
              <p class="meta">${a.type || 'Activity'} • ${fmtDistanceMeters(a.distance)} • ${fmtHours(a.moving_time)} • ${activityToTss(a)} TSS</p>
            </div>
            <span class="badge done">Done</span>
          `;
          row.addEventListener('click', () => openWorkoutModal({ source: 'strava', data: a }));
          row.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'strava', data: a }));
          doneNode.appendChild(row);
        });
      }

      const plannedNode = document.getElementById('todayPlanned');
      plannedNode.innerHTML = '';
      if (!plannedUpcoming.length) {
        plannedNode.innerHTML = '<p class="meta">No planned workouts yet. Use + on any calendar day.</p>';
      } else {
        plannedUpcoming.forEach(p => {
          const pair = pairForPlanned(String(p.id));
          const pairedCompleted = pair ? activities.find(a => String(a.id) === String(pair.strava_id)) : null;
          const row = document.createElement('button');
          row.type = 'button';
          row.className = 'list-item';
          row.innerHTML = `
            <div>
              <p class="title">${p.title || p.workout_type}</p>
              <p class="meta">${p.date} • ${p.workout_type} • ${Number(p.duration_min || 0)} min • ${fmtDistanceKm(p.distance_km)} • ${itemToTss(p)} TSS</p>
            </div>
            <span class="badge planned">Planned</span>
          `;
          row.addEventListener('click', () => openWorkoutModal({ source: pairedCompleted ? 'strava' : 'planned', data: pairedCompleted || p, planned: p, pair }));
          row.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'planned', data: p }));
          plannedNode.appendChild(row);
        });
      }

      renderEvents();
      renderGoals();
      renderPerformanceMetrics();
    }

    function renderCalendar(options = {}) {
      const { forceAnchor = false, preserveScroll = true, targetDateKey = '' } = options;
      const dayMap = buildDayAggregateMap();
      const plannedById = new Map(calendarItems.filter(i => i.kind === 'workout').map(i => [String(i.id), i]));
      const stravaById = new Map(activities.map(a => [String(a.id), a]));
      const pairByPlannedId = new Map(pairs.map(p => [String(p.planned_id), p]));
      const pairByStravaId = new Map(pairs.map(p => [String(p.strava_id), p]));
      const wrap = document.getElementById('calendarScroll');
      const prevScroll = preserveScroll ? wrap.scrollTop : 0;
      const priorRows = preserveScroll ? Array.from(wrap.querySelectorAll('.week-row')) : [];
      const wrapTop = wrap.getBoundingClientRect().top;
      let priorTopWeek = null;
      for (const row of priorRows) {
        if (row.getBoundingClientRect().bottom >= wrapTop + 2) {
          priorTopWeek = row;
          break;
        }
      }
      const preservedWeekStart = priorTopWeek ? String(priorTopWeek.dataset.weekStart || '') : '';
      wrap.innerHTML = '';
      const grid = document.createElement('section');
      grid.className = 'month';

      const baseDate = targetDateKey ? parseDateKey(targetDateKey) : new Date(calendarCursor.getFullYear(), calendarCursor.getMonth(), 1);
      const startMonth = new Date(baseDate.getFullYear(), baseDate.getMonth() - 4, 1);
      const endMonth = new Date(baseDate.getFullYear(), baseDate.getMonth() + 9, 0);
      const startDate = new Date(startMonth);
      const startOffset = (startDate.getDay() + 6) % 7;
      startDate.setDate(startDate.getDate() - startOffset);
      const endDate = new Date(endMonth);
      const endOffset = (7 - ((endDate.getDay() + 6) % 7) - 1);
      endDate.setDate(endDate.getDate() + endOffset);

      const weekRows = [];
      for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 7)) {
        const row = document.createElement('div');
        row.className = 'week-row';
        const weekDateKeys = [];
        const weekStart = new Date(d);
        const weekMid = new Date(d);
        weekMid.setDate(weekMid.getDate() + 3);
        row.dataset.weekMonth = monthKey(weekMid.getFullYear(), weekMid.getMonth());
        row.dataset.weekLabel = weekMid.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
        row.dataset.weekStart = dateKeyFromDate(weekStart);

        for (let col = 0; col < 7; col += 1) {
          const dayDate = new Date(weekStart);
          dayDate.setDate(weekStart.getDate() + col);
          const key = dateKeyFromDate(dayDate);
          weekDateKeys.push(key);

          const cell = document.createElement('div');
          cell.className = 'day';
          cell.dataset.date = key;
          if (key === todayKey()) cell.classList.add('today');
          if (dayDate.getMonth() !== calendarCursor.getMonth()) cell.style.opacity = '0.92';

          const num = document.createElement('span');
          num.className = 'd-num';
          num.textContent = String(dayDate.getDate());
          cell.appendChild(num);

          const entries = dayMap[key] || { done: [], items: [] };
          const shownCompleted = new Set();
          const cardsToShow = [];
          entries.items.forEach((item) => {
            if (item.kind !== 'workout') {
              cardsToShow.push({ kind: 'other', item });
              return;
            }
            const pair = pairByPlannedId.get(String(item.id));
            const completed = pair ? stravaById.get(String(pair.strava_id)) : completedFromPlanned(item);
            if (completed) shownCompleted.add(String(completed.id));
            cardsToShow.push({ kind: 'planned', item, completed, pair, fromPair: !!pair });
          });
          entries.done.forEach((a) => {
            if (!shownCompleted.has(String(a.id))) {
              const pair = pairByStravaId.get(String(a.id));
              cardsToShow.push({ kind: 'completed', completed: a, pair });
            }
          });

          cardsToShow.slice(0, 6).forEach((entry) => {
            if (entry.kind === 'other') {
              const item = entry.item;
              const card = document.createElement('div');
              card.className = `work-card ${item.kind}`;
              card.innerHTML = `<button class="card-menu-btn" type="button">&#8942;</button><p class="wc-title">${item.title || item.kind.toUpperCase()}</p><p class="wc-meta">${item.kind.toUpperCase()} • ${item.date}</p>`;
              card.addEventListener('click', (ev) => {
                ev.stopPropagation();
                selectedKind = item.kind;
                selectedDate = item.date;
                selectedWorkoutType = item.workout_type || 'Other';
                openDetailModal(item);
              });
              card.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
              card.querySelector('.card-menu-btn').addEventListener('click', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
              cell.appendChild(card);
              return;
            }
            if (entry.kind === 'planned') {
              const item = entry.item;
              const completed = entry.completed;
              const comp = complianceStatus(item, completed, key);
              const card = document.createElement('div');
              card.className = `work-card ${comp.cls}`;
              const unitForCard = String(item.distance_unit || distanceUnit || 'km');
              const plannedDistM = Number(item.distance_m || (Number(item.distance_km || 0) * 1000));
              const durMin = completed ? Number(completed.moving_time || 0) / 60 : Number(item.duration_min || 0);
              const distM = completed ? Number(completed.distance || 0) : plannedDistM;
              const tssVal = completed
                ? Math.round(Number(completed.tss_override || 0) || activityToTss(completed))
                : Math.round(Number(item.planned_tss || 0));
              const cTime = completed ? formatStartClock(completed.start_date_local) : '';
              const arrow = comp.arrow === 'up' ? '<span class="delta-up">↑</span>' : comp.arrow === 'down' ? '<span class="delta-down">↓</span>' : '';
              const feedCount = commentCount(item) || commentCount(completed);
              const feedback = `${feelEmoji(item.feel)} ${Number(item.rpe || 0) > 0 ? item.rpe : ''}`.trim();
              const commentsText = feedCount > 0 ? `💬 x${feedCount}` : '💬';
              const metricParts = [];
              if (durMin > 0) metricParts.push(`<strong>${formatDurationClockCompact(durMin)}</strong>`);
              if (distM > 0) metricParts.push(`<strong>${fmtDistanceMetersInUnit(distM, unitForCard)}</strong>`);
              if (tssVal > 0) metricParts.push(`<strong>${tssVal} TSS</strong>`);
              card.innerHTML = `<button class="card-menu-btn" type="button">&#8942;</button>${cardIcon(item.workout_type)}<p class="wc-title"><span>${item.title || (item.workout_type || 'Workout')}</span></p>${metricParts.length ? `<p class="wc-metrics">${metricParts.join('')}</p>` : ''}${cTime ? `<p class="wc-meta">C: ${cTime} ${arrow}</p>` : ''}<div class="wc-bottom"><span>${feedback || '&nbsp;'}</span><span>${commentsText}</span></div>`;
              card.draggable = true;
              card.dataset.kind = 'planned';
              card.dataset.plannedId = String(item.id);
              card.addEventListener('dragstart', (ev) => ev.dataTransfer.setData('text/plain', JSON.stringify({ source: 'planned', id: String(item.id) })));
              card.addEventListener('dragover', (ev) => ev.preventDefault());
              card.addEventListener('drop', async (ev) => {
                ev.preventDefault();
                const raw = ev.dataTransfer.getData('text/plain');
                if (!raw) return;
                const dragData = JSON.parse(raw);
                if (dragData.source === 'strava') await pairWorkouts(String(item.id), String(dragData.id));
              });
              card.addEventListener('click', (ev) => {
                ev.stopPropagation();
                const source = entry.fromPair && completed ? 'strava' : 'planned';
                openWorkoutModal({ source, data: source === 'strava' ? completed : item, planned: item, pair: entry.pair });
              });
              card.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
              card.querySelector('.card-menu-btn').addEventListener('click', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
              cell.appendChild(card);
              return;
            }
            const a = entry.completed;
            const pairedPlanned = entry.pair ? plannedById.get(String(entry.pair.planned_id)) : null;
            const compStat = pairedPlanned ? complianceStatus(pairedPlanned, a, key).cls : 'unplanned';
            const unitForCard = String((pairedPlanned && pairedPlanned.distance_unit) || a.distance_unit || distanceUnit || 'km');
            const card = document.createElement('div');
            card.className = `work-card ${compStat}`;
            const feedCount = commentCount(a);
            const feedback = `${feelEmoji(a.feel)} ${Number(a.rpe || 0) > 0 ? a.rpe : ''}`.trim();
            const commentsText = feedCount > 0 ? `💬 x${feedCount}` : '💬';
            const cDur = Number(a.moving_time || 0) / 60;
            const cDist = Number(a.distance || 0);
            const cTss = Math.round(activityToTss(a));
            const metricParts = [];
            if (cDur > 0) metricParts.push(`<strong>${formatDurationClockCompact(cDur)}</strong>`);
            if (cDist > 0) metricParts.push(`<strong>${fmtDistanceMetersInUnit(cDist, unitForCard)}</strong>`);
            if (cTss > 0) metricParts.push(`<strong>${cTss} TSS</strong>`);
            card.innerHTML = `<button class="card-menu-btn" type="button">&#8942;</button>${cardIcon(a.type)}<p class="wc-title"><span>${a.name || 'Completed Workout'}</span></p>${metricParts.length ? `<p class="wc-metrics">${metricParts.join('')}</p>` : ''}<p class="wc-meta">C: ${formatStartClock(a.start_date_local)}</p><div class="wc-bottom"><span>${feedback || '&nbsp;'}</span><span>${commentsText}</span></div>`;
            card.draggable = true;
            card.dataset.kind = 'strava';
            card.dataset.stravaId = String(a.id);
            card.addEventListener('dragstart', (ev) => ev.dataTransfer.setData('text/plain', JSON.stringify({ source: 'strava', id: String(a.id) })));
            card.addEventListener('dragover', (ev) => ev.preventDefault());
            card.addEventListener('drop', async (ev) => {
              ev.preventDefault();
              const raw = ev.dataTransfer.getData('text/plain');
              if (!raw) return;
              const dragData = JSON.parse(raw);
              if (dragData.source === 'planned') await pairWorkouts(String(dragData.id), String(a.id));
            });
            card.addEventListener('click', (ev) => {
              ev.stopPropagation();
              openWorkoutModal({ source: 'strava', data: a, pair: entry.pair });
            });
            card.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'strava', data: a }));
            card.querySelector('.card-menu-btn').addEventListener('click', (ev) => showItemMenu(ev, { source: 'strava', data: a }));
            cell.appendChild(card);
          });

          if (cardsToShow.length > 6) {
            const more = document.createElement('span');
            more.className = 'item';
            more.style.background = '#edf3fb';
            more.style.color = '#5c7898';
            more.textContent = `+${cardsToShow.length - 6} more`;
            cell.appendChild(more);
          }

          const addBar = document.createElement('button');
          addBar.type = 'button';
          addBar.className = 'quick-add';
          addBar.textContent = '+';
          addBar.title = 'Add item';
          addBar.addEventListener('click', (ev) => {
            ev.stopPropagation();
            openActionModal(key);
          });
          cell.appendChild(addBar);
          row.appendChild(cell);
        }

        const week = getWeekMetrics(weekDateKeys, dayMap);
        const weekCard = document.createElement('div');
        weekCard.className = 'week-summary';
        weekCard.innerHTML = `<div class="ws-metrics"><div class="ws-chip ws-ctl"><strong>${week.ctl}</strong>CTL</div><div class="ws-chip ws-atl"><strong>${week.atl}</strong>ATL</div><div class="ws-chip ws-tsb"><strong>${week.tsb > 0 ? '+' + week.tsb : week.tsb}</strong>TSB</div></div><div class="ws-row"><span>Total Duration</span><strong>${week.durationLabel}</strong></div><div class="ws-row"><span>Total TSS</span><strong>${week.tss}</strong></div>`;
        row.appendChild(weekCard);
        grid.appendChild(row);
        weekRows.push(row);
      }

      wrap.appendChild(grid);

      if (!calendarScrollBound) {
        wrap.addEventListener('scroll', syncCalendarHeaderFromScroll);
        calendarScrollBound = true;
      }
      const currentStartKey = dateKeyFromDate(new Date(baseDate.getFullYear(), baseDate.getMonth(), 1));
      let rowTarget = weekRows[0] || null;
      weekRows.forEach((row) => {
        if (String(row.dataset.weekStart || '') <= currentStartKey) rowTarget = row;
      });
      if (targetDateKey) {
        const targetCell = wrap.querySelector(`.day[data-date="${targetDateKey}"]`);
        const targetRow = targetCell ? targetCell.closest('.week-row') : null;
        if (targetRow) rowTarget = targetRow;
      }
      if (!forceAnchor) {
        const preferred = preservedWeekStart || calendarAnchorWeekStart;
        if (preferred) {
          const preferredRow = weekRows.find((row) => String(row.dataset.weekStart || '') === preferred);
          if (preferredRow) rowTarget = preferredRow;
        }
      }
      if (forceAnchor && rowTarget) {
        wrap.scrollTop = rowTarget.offsetTop;
        calendarAnchorWeekStart = String(rowTarget.dataset.weekStart || calendarAnchorWeekStart);
      } else if (preserveScroll && prevScroll > 0) {
        wrap.scrollTop = calendarScrollTop > 0 ? calendarScrollTop : prevScroll;
      } else if (rowTarget) {
        wrap.scrollTop = rowTarget.offsetTop;
      }
      syncCalendarHeaderFromScroll();
    }

    function jumpToCurrentMonth() {
      const now = new Date();
      calendarCursor = new Date(now.getFullYear(), now.getMonth(), 1);
      calendarAnchorWeekStart = dateKeyFromDate(mondayOfDate(now));
      calendarScrollTop = 0;
      renderCalendar({ forceAnchor: true, preserveScroll: false, targetDateKey: todayKey() });
    }

    function renderDashboard() {
      const totalDistance = activities.reduce((sum, a) => sum + Number(a.distance || 0), 0);
      const totalTime = activities.reduce((sum, a) => sum + Number(a.moving_time || 0), 0);
      document.getElementById('statCount').textContent = String(activities.length);
      document.getElementById('statPlanned').textContent = String(calendarItems.filter(i => i.kind === 'workout').length);
      document.getElementById('statDistance').textContent = fmtDistanceMeters(totalDistance);
      document.getElementById('statTime').textContent = fmtHours(totalTime);
    }

    function renderSettings() {
      const ftp = appSettings.ftp || {};
      const setVal = (id, key) => {
        const el = document.getElementById(id);
        if (!el) return;
        const v = ftp[key];
        el.value = v == null ? '' : String(v);
      };
      setVal('ftpRide', 'ride');
      setVal('ftpRun', 'run');
      setVal('ftpRow', 'row');
      setVal('ftpSwim', 'swim');
      setVal('ftpStrength', 'strength');
      setVal('ftpOther', 'other');
    }

    function applyWidgetPrefs() {
      const pref = JSON.parse(localStorage.getItem('dashboardWidgets') || '{}');
      ['count', 'plannedCount', 'distance', 'time'].forEach(key => {
        const visible = pref[key] !== false;
        const card = document.querySelector(`[data-widget="${key}"]`);
        const toggle = document.querySelector(`[data-toggle="${key}"]`);
        if (card) card.style.display = visible ? 'block' : 'none';
        if (toggle) toggle.checked = visible;
      });
    }

    function bindWidgetToggles() {
      document.querySelectorAll('[data-toggle]').forEach(input => {
        input.addEventListener('change', () => {
          const key = input.getAttribute('data-toggle');
          const pref = JSON.parse(localStorage.getItem('dashboardWidgets') || '{}');
          pref[key] = input.checked;
          localStorage.setItem('dashboardWidgets', JSON.stringify(pref));
          applyWidgetPrefs();
        });
      });
    }

    async function loadData(resetMonthPosition) {
      try {
        const [aResp, cResp, pResp, sResp] = await Promise.all([fetch('/ui/activities'), fetch('/calendar-items'), fetch('/pairs'), fetch('/settings')]);
        activities = aResp.ok ? await aResp.json() : [];
        calendarItems = cResp.ok ? await cResp.json() : [];
        pairs = pResp.ok ? await pResp.json() : [];
        appSettings = sResp.ok ? await sResp.json() : { units: { distance: 'km', elevation: 'm' }, ftp: {} };
        if (appSettings.units && appSettings.units.distance) {
          distanceUnit = appSettings.units.distance;
        }
        if (appSettings.units && appSettings.units.elevation) {
          elevationUnit = appSettings.units.elevation;
        }
      } catch (_err) {
        activities = [];
        calendarItems = [];
        pairs = [];
      }

      if (resetMonthPosition) initialMonthCentered = false;
      updateUnitButtons();
      renderHome();
      if (isCalendarActive()) {
        renderCalendar({ forceAnchor: resetMonthPosition, preserveScroll: !resetMonthPosition });
      }
      renderDashboard();
      renderSettings();
    }

    document.querySelectorAll('.tab').forEach(btn => {
      btn.addEventListener('click', () => setView(btn.dataset.view));
    });

    document.getElementById('saveSettingsBtn').addEventListener('click', async () => {
      const read = (id) => {
        const v = document.getElementById(id).value.trim();
        return v ? Number(v) : null;
      };
      const payload = {
        units: {
          distance: (appSettings.units && appSettings.units.distance) || distanceUnit || 'km',
          elevation: (appSettings.units && appSettings.units.elevation) || elevationUnit || 'm',
        },
        ftp: {
          ride: read('ftpRide'),
          run: read('ftpRun'),
          row: read('ftpRow'),
          swim: read('ftpSwim'),
          strength: read('ftpStrength'),
          other: read('ftpOther'),
        },
      };
      const resp = await fetch('/settings', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      if (!resp.ok) return;
      appSettings = await resp.json();
      const msg = document.getElementById('settingsSavedMsg');
      msg.style.display = 'inline';
      setTimeout(() => { msg.style.display = 'none'; }, 1400);
      await loadData(false);
    });
    document.getElementById('uploadFitBtn').addEventListener('click', () => {
      fitUploadContext = 'global';
      fitUploadTargetActivityId = null;
      document.getElementById('uploadFitInput').click();
    });
    document.getElementById('uploadFitInput').addEventListener('change', async (event) => {
      const input = event.target;
      const file = input.files && input.files[0];
      if (!file) return;
      input.value = '';
      if (fitUploadContext === 'modal') {
        const payload = window.currentWorkoutPayload;
        if (!payload || !modalDraft) return;
        const targetPlanned = payload.planned || (payload.source === 'planned' ? payload.data : null);
        if (payload.source === 'strava') {
          const resp = await fetch(`/activities/${encodeURIComponent(payload.data.id)}/fit/upload?filename=${encodeURIComponent(file.name)}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/octet-stream' },
            body: file,
          });
          if (resp.ok) {
            const uploaded = await resp.json();
            modalDraft.uploadedNow = true;
            window.currentWorkoutPayload = { ...payload, data: uploaded };
            renderWorkoutSummary(window.currentWorkoutPayload);
            renderWorkoutFiles(window.currentWorkoutPayload);
            await renderWorkoutAnalyze(window.currentWorkoutPayload);
          }
        } else if (targetPlanned) {
          const upResp = await fetch(`/import-fit?filename=${encodeURIComponent(file.name)}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/octet-stream' },
            body: file,
          });
          if (upResp.ok) {
            const uploaded = await upResp.json();
            const pairResp = await fetch('/pairs', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                planned_id: targetPlanned.id,
                strava_id: String(uploaded.id),
                override_date: targetPlanned.date,
                override_title: targetPlanned.title || 'Untitled Workout',
              }),
            });
            const pair = pairResp.ok ? await pairResp.json() : null;
            modalDraft.uploadedNow = true;
            modalDraft.createdActivityId = uploaded.id;
            modalDraft.createdPairId = pair ? pair.id : null;
            window.currentWorkoutPayload = { source: 'strava', data: uploaded, planned: targetPlanned, pair };
            renderWorkoutSummary(window.currentWorkoutPayload);
            renderWorkoutFiles(window.currentWorkoutPayload);
            await renderWorkoutAnalyze(window.currentWorkoutPayload);
          }
        }
        fitUploadTargetActivityId = null;
        fitUploadContext = 'global';
        return;
      }
      const endpoint = fitUploadTargetActivityId
        ? `/activities/${encodeURIComponent(fitUploadTargetActivityId)}/fit/upload?filename=${encodeURIComponent(file.name)}`
        : `/import-fit?filename=${encodeURIComponent(file.name)}`;
      const resp = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream' },
        body: file,
      });
      if (!resp.ok) {
        const err = await resp.text();
        alert(`Import failed: ${err}`);
        fitUploadTargetActivityId = null;
        fitUploadContext = 'global';
        return;
      }
      const uploaded = await resp.json();
      if (fitUploadTargetActivityId && window.currentWorkoutPayload && window.currentWorkoutPayload.data) {
        window.currentWorkoutPayload = { ...window.currentWorkoutPayload, data: uploaded };
        renderWorkoutFiles(window.currentWorkoutPayload);
        renderWorkoutSummary(window.currentWorkoutPayload);
        await renderWorkoutAnalyze(window.currentWorkoutPayload);
      }
      fitUploadTargetActivityId = null;
      fitUploadContext = 'global';
      await loadData(false);
    });
    document.getElementById('globalSettings').addEventListener('click', () => {
      document.getElementById('settingsModal').classList.add('open');
    });
    document.getElementById('closeSettings').addEventListener('click', () => {
      document.getElementById('settingsModal').classList.remove('open');
    });
    document.getElementById('settingsModal').addEventListener('click', (event) => {
      if (event.target.id === 'settingsModal') document.getElementById('settingsModal').classList.remove('open');
    });
    document.getElementById('addEventBtn').addEventListener('click', () => openActionModal(todayKey(), 'event'));
    document.getElementById('addGoalBtn').addEventListener('click', () => openActionModal(todayKey(), 'goal'));

    document.getElementById('closeAction').addEventListener('click', closeActionModal);
    document.getElementById('actionModal').addEventListener('click', (event) => {
      if (event.target.id === 'actionModal') closeActionModal();
    });

    document.getElementById('cancelDetail').addEventListener('click', closeDetailModal);
    document.getElementById('deleteDetail').addEventListener('click', deleteCurrentDetail);
    document.getElementById('saveDetail').addEventListener('click', () => saveDetail(false));
    document.getElementById('saveCloseDetail').addEventListener('click', () => saveDetail(true));
    document.getElementById('detailModal').addEventListener('click', (event) => {
      if (event.target.id === 'detailModal') closeDetailModal();
    });

    document.getElementById('closeWorkoutView').addEventListener('click', async () => { await closeWorkoutModal(true); await loadData(false); });
    document.getElementById('cancelWorkoutView').addEventListener('click', async () => { await closeWorkoutModal(true); await loadData(false); });
    document.getElementById('saveWorkoutView').addEventListener('click', () => saveWorkoutView(false));
    document.getElementById('saveCloseWorkoutView').addEventListener('click', () => saveWorkoutView(true));
    document.getElementById('deleteWorkoutView').addEventListener('click', async () => {
      const payload = window.currentWorkoutPayload;
      if (!payload) return;
      if (!window.confirm('Are you sure you want to delete this?')) return;
      const data = payload.data || {};
      if (payload.planned || payload.source === 'planned') {
        const targetId = payload.planned ? payload.planned.id : data.id;
        if (!targetId) {
          await closeWorkoutModal(false);
          return;
        }
        await fetch(`/calendar-items/${targetId}`, { method: 'DELETE' });
      } else if (payload.source === 'strava') {
        await fetch(`/activities/${data.id}`, { method: 'DELETE' });
      }
      await closeWorkoutModal(false);
      await loadData(false);
    });
    document.getElementById('wvBrowseFilesBtn').addEventListener('click', () => {
      const payload = window.currentWorkoutPayload;
      if (!payload) return;
      fitUploadContext = 'modal';
      fitUploadTargetActivityId = payload.source === 'strava' ? payload.data.id : null;
      document.getElementById('uploadFitInput').click();
    });
    document.getElementById('wvFilesTabBtn').addEventListener('click', () => {
      const pop = document.getElementById('wvFilesPopover');
      const isHidden = pop.classList.contains('hidden');
      pop.classList.toggle('hidden', !isHidden);
      document.getElementById('wvFilesTabBtn').classList.toggle('active', isHidden);
    });
    document.getElementById('wvAnalyzeBtn').addEventListener('click', () => {
      const showingAnalyze = !document.getElementById('wvAnalyze').classList.contains('hidden');
      setWorkoutMode(showingAnalyze ? 'summary' : 'analyze');
    });
    document.getElementById('wvSportToggle').addEventListener('click', (ev) => {
      ev.stopPropagation();
      document.getElementById('wvSportMenu').classList.toggle('hidden');
    });
    document.getElementById('wvSportIcon').addEventListener('click', (ev) => {
      ev.stopPropagation();
      document.getElementById('wvSportMenu').classList.toggle('hidden');
    });
    document.querySelectorAll('.feel-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (btn.disabled) return;
        setFeelValue(btn.dataset.feel);
      });
    });
    document.getElementById('wvRpe').addEventListener('input', () => {
      if (modalDraft) modalDraft.rpeTouched = true;
      const rpe = document.getElementById('wvRpe');
      rpe.classList.remove('rpe-unset');
      updateRpeLabel();
    });
    document.getElementById('wvPostCommentBtn').addEventListener('click', () => {
      const input = document.getElementById('wvCommentInput');
      const txt = String(input.value || '').trim();
      if (!txt || !modalDraft) return;
      modalDraft.commentsFeed = modalDraft.commentsFeed || [];
      modalDraft.commentsFeed.push(txt);
      input.value = '';
      renderCommentsFeed();
    });
    document.getElementById('wvCommentsFeed').addEventListener('click', (ev) => {
      const del = ev.target.closest('.comment-delete');
      if (!del || !modalDraft) return;
      const item = del.closest('.comment-item');
      if (!item) return;
      const idx = Number(item.dataset.index);
      if (!Number.isInteger(idx) || idx < 0) return;
      modalDraft.commentsFeed.splice(idx, 1);
      renderCommentsFeed();
    });
    document.getElementById('wvCommentsFeed').addEventListener('dblclick', (ev) => {
      const textNode = ev.target.closest('.comment-text');
      if (!textNode || !modalDraft) return;
      const item = textNode.closest('.comment-item');
      if (!item) return;
      const idx = Number(item.dataset.index);
      if (!Number.isInteger(idx) || idx < 0) return;
      const input = document.createElement('textarea');
      input.value = textNode.textContent || '';
      input.className = 'comment-edit-input';
      const delBtn = item.querySelector('.comment-delete');
      if (delBtn) delBtn.textContent = '×';
      textNode.replaceWith(input);
      input.focus();
      input.select();
      const commit = () => {
        modalDraft.commentsFeed[idx] = String(input.value || '').trim();
        modalDraft.commentsFeed = modalDraft.commentsFeed.filter(Boolean);
        renderCommentsFeed();
      };
      input.addEventListener('blur', commit, { once: true });
      input.addEventListener('keydown', (keyEv) => {
        if (keyEv.key === 'Enter' && !keyEv.shiftKey) {
          keyEv.preventDefault();
          commit();
        }
      });
    });
    document.getElementById('wvCommentInput').addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter') {
        if (ev.shiftKey) return;
        ev.preventDefault();
        document.getElementById('wvPostCommentBtn').click();
      }
    });
    document.getElementById('workoutViewModal').addEventListener('click', (event) => {
      if (event.target.id === 'workoutViewModal') {
        closeWorkoutModal(true).then(() => loadData(false));
      }
    });
    document.getElementById('contextMenu').addEventListener('click', (ev) => ev.stopPropagation());
    document.addEventListener('click', (ev) => {
      const sportMenu = document.getElementById('wvSportMenu');
      if (sportMenu && !sportMenu.classList.contains('hidden') && !ev.target.closest('#wvSportMenu') && !ev.target.closest('#wvSportToggle')) {
        sportMenu.classList.add('hidden');
      }
      if (!ev.target.closest('.card-menu-btn') && !ev.target.closest('#contextMenu')) {
        closeContextMenu();
      }
    });
    document.getElementById('calTodayBtn').addEventListener('click', jumpToCurrentMonth);

    buildTypeGrids();
    bindWidgetToggles();
    applyWidgetPrefs();
    updateUnitButtons();
    setView('home');
    loadData(true);
