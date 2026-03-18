    // ── Custom delete confirm modal ──
    let _deleteConfirmResolver = null;
    let _applyConfirmResolver = null;
    let _commentDeleteResolver = null;
    let _unsavedCloseResolver = null;
    function confirmDelete({ message = 'This cannot be undone.', onConfirm } = {}) {
      const modal = document.getElementById('deleteConfirmModal');
      const subtitle = document.getElementById('deleteConfirmSub');
      if (subtitle) subtitle.textContent = message;
      modal.classList.add('open');
      return new Promise((resolve) => {
        _deleteConfirmResolver = async (ok) => {
          resolve(Boolean(ok));
          if (ok && typeof onConfirm === 'function') await onConfirm();
        };
      });
    }

    function confirmApplyChanges() {
      const modal = document.getElementById('applyConfirmModal');
      if (!modal) return Promise.resolve(false);
      modal.classList.add('open');
      return new Promise((resolve) => {
        _applyConfirmResolver = (ok) => {
          resolve(Boolean(ok));
        };
      });
    }

    function confirmDeleteComment() {
      const modal = document.getElementById('commentDeleteConfirmModal');
      if (!modal) return Promise.resolve(false);
      modal.classList.add('open');
      return new Promise((resolve) => {
        _commentDeleteResolver = (ok) => {
          resolve(Boolean(ok));
        };
      });
    }

    function confirmUnsavedClose() {
      const modal = document.getElementById('unsavedCloseConfirmModal');
      if (!modal) return Promise.resolve('cancel');
      modal.classList.add('open');
      return new Promise((resolve) => {
        _unsavedCloseResolver = (choice) => resolve(choice || 'cancel');
      });
    }

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
      ride: '/icons/workouts/bike.png',
      bike: '/icons/workouts/bike.png',
      swim: '/icons/workouts/swim.png',
      brick: '/icons/workouts/brick.png',
      pulse: '/icons/workouts/crosstrain.png',
      rest: '/icons/workouts/day_off.png',
      mtb: '/icons/workouts/mountian_bike.png',
      strength: '/icons/workouts/strength.png',
      timer: '/icons/workouts/other_custom.png',
      ski: '/icons/workouts/XC_Ski.png',
      row: '/icons/workouts/row.png',
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
    let currentDragData = null; // tracks active drag payload reliably
    let selectedDate = todayKey();
    let selectedKind = 'workout';
    let selectedWorkoutType = 'Run';
    let editingItemId = null;
    let analyzeState = null;
    let appSettings = { units: { distance: 'km', elevation: 'm' }, ftp: {} };
    let distanceUnit = localStorage.getItem('distanceUnit') || 'km';
    let elevationUnit = localStorage.getItem('elevationUnit') || 'm';
    let currentFeel = 0;
    let fitUploadTargetActivityId = null;
    let fitUploadContext = 'global';
    let modalDraft = null;
    let detailInitialState = null;
    let workoutModalSession = 0;
    const calendarState = {
      anchorDate: todayKey(),
      scrollTop: 0,
      viewportAnchor: null,
      hasRendered: false,
      scrollDebounce: null,
      activeScrollSync: false,
    };

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

    function getCalendarScrollContainer() {
      return document.getElementById('calendarScroll');
    }

    function rememberCalendarPosition() {
      const wrap = getCalendarScrollContainer();
      if (!wrap) return;
      calendarState.scrollTop = wrap.scrollTop;
      calendarState.viewportAnchor = null;
      const containerRect = wrap.getBoundingClientRect();
      const top = containerRect.top;
      const bottom = containerRect.bottom;

      const days = Array.from(wrap.querySelectorAll('.day[data-date]'));
      const dayAnchor = days.find((el) => {
        const r = el.getBoundingClientRect();
        return r.bottom >= top && r.top <= bottom;
      });
      if (dayAnchor && dayAnchor.dataset.date) {
        const r = dayAnchor.getBoundingClientRect();
        calendarState.viewportAnchor = {
          type: 'day',
          key: String(dayAnchor.dataset.date),
          offsetFromTop: r.top - top,
        };
        calendarState.anchorDate = String(dayAnchor.dataset.date);
        return;
      }

      const rows = Array.from(wrap.querySelectorAll('.week-row[data-week-start]'));
      const rowAnchor = rows.find((el) => {
        const r = el.getBoundingClientRect();
        return r.bottom >= top && r.top <= bottom;
      });
      if (rowAnchor && rowAnchor.dataset.weekStart) {
        const r = rowAnchor.getBoundingClientRect();
        calendarState.viewportAnchor = {
          type: 'week',
          key: String(rowAnchor.dataset.weekStart),
          offsetFromTop: r.top - top,
        };
      }
    }

    function restoreCalendarPositionFromAnchor() {
      const wrap = getCalendarScrollContainer();
      const anchor = calendarState.viewportAnchor;
      if (!wrap || !anchor || !anchor.type || !anchor.key) return false;
      const selector = anchor.type === 'day'
        ? `.day[data-date="${anchor.key}"]`
        : `.week-row[data-week-start="${anchor.key}"]`;
      const target = wrap.querySelector(selector);
      if (!target) return false;
      const containerRect = wrap.getBoundingClientRect();
      const targetRect = target.getBoundingClientRect();
      const delta = (targetRect.top - containerRect.top) - Number(anchor.offsetFromTop || 0);
      wrap.scrollTop += delta;
      calendarState.scrollTop = wrap.scrollTop;
      return true;
    }

    function syncCalendarHeaderFromScroll() {
      const wrap = getCalendarScrollContainer();
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
      const topDay = row.querySelector('.day[data-date]');
      if (topDay && topDay.dataset.date) {
        calendarState.anchorDate = String(topDay.dataset.date);
      }
      calendarState.scrollTop = wrap.scrollTop;
    }

    function bindCalendarScrollSync() {
      const wrap = getCalendarScrollContainer();
      if (!wrap || calendarState.activeScrollSync) return;
      wrap.addEventListener('scroll', () => {
        calendarState.scrollTop = wrap.scrollTop;
        if (calendarState.scrollDebounce) clearTimeout(calendarState.scrollDebounce);
        calendarState.scrollDebounce = setTimeout(() => {
          syncCalendarHeaderFromScroll();
        }, 50);
      });
      calendarState.activeScrollSync = true;
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

    function isFutureDateKey(key) {
      if (!key) return false;
      return key > todayKey();
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
      if (pageHead) pageHead.classList.toggle('hidden', name === 'calendar' || name === 'home');
      document.getElementById('pageTitle').textContent = name.charAt(0).toUpperCase() + name.slice(1);
      if (name === 'calendar') {
        if (!calendarState.hasRendered) {
          renderCalendar({ preserveScroll: false, anchorDate: calendarState.anchorDate, jumpToDate: calendarState.anchorDate });
        } else {
          syncCalendarHeaderFromScroll();
        }
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

    function lthrForActivity(activity) {
      const lthr = appSettings.lthr || {};
      const key = activitySportKey(activity);
      const v = Number(lthr[key] || lthr.global || 0);
      return v > 0 ? v : null;
    }

    // Coggan zone upper-bound percentages of LTHR (Z1–Z4; Z5 = above last)
    const HR_ZONE_BOUNDS = [68, 84, 95, 106];
    const HR_ZONE_RATES  = [30, 55, 70, 90, 110]; // TSS/hr per zone

    function calcHrTss(activity) {
      // Use pre-calculated value from FIT parsing when available
      const preCalc = Number(activity.hr_tss || 0);
      if (preCalc > 0) return preCalc;
      // Estimate from avg HR + duration (Strava activities without FIT)
      const lthr = lthrForActivity(activity);
      if (!lthr) return null;
      const avgHr = Number(activity.avg_heartrate || activity.avg_hr || 0);
      if (!avgHr) return null;
      const durationH = Number(activity.moving_time || 0) / 3600;
      if (durationH <= 0) return null;
      const pct = (avgHr / lthr) * 100;
      const zoneIdx = HR_ZONE_BOUNDS.findIndex(b => pct < b);
      const rate = HR_ZONE_RATES[zoneIdx === -1 ? 4 : zoneIdx];
      return durationH * rate;
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
      const powerTss = (ifv && durationH > 0) ? durationH * ifv * ifv * 100 : null;
      const hrTss = calcHrTss(activity);
      if (activity.tss_source === 'hr' && hrTss) return hrTss;
      if (activity.tss_source === 'power' && powerTss) return powerTss;
      if (powerTss) return powerTss;
      if (hrTss) return hrTss;
      return estimateTss(Number(activity.moving_time || 0) / 60, intensityByType(activity.type || 'Other'));
    }

    function completedDurationMinValue(obj) {
      const overrideMin = Number(obj && obj.completed_duration_min);
      if (Number.isFinite(overrideMin) && overrideMin >= 0) return overrideMin;
      return Number(obj && obj.moving_time || 0) / 60;
    }

    function itemToTss(item) {
      if (item.kind !== 'workout') return 0;
      const plannedTss = Number(item.planned_tss || 0);
      if (plannedTss > 0) return plannedTss;
      const userIntensity = Number(item.intensity || 0);
      const intensity = userIntensity > 0 ? (0.45 + (Math.min(10, userIntensity) / 10) * 0.55) : intensityByType(item.workout_type || 'Other');
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
      const hours = completedDurationMinValue(obj) / 60;
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

    function hasPlannedAndCompletedContent(workout) {
      if (!workout || workout.kind !== 'workout') return false;
      const hasPlanned = ['duration', 'distance', 'tss'].some((basis) => plannedMetric(workout, basis) > 0);
      const hasCompleted = !!completedFromPlanned(workout);
      return hasPlanned && hasCompleted;
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
        const h = completedDurationMinValue(completedItem) / 60;
        if (ifv > 0 && h > 0) return h * ifv * ifv * 100;
        return activityToTss(completedItem);
      }
      return completedDurationMinValue(completedItem);
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

    function modalStatusClass(payload) {
      if (!payload) return 'status-gray';
      if (payload.isDraft) return 'status-gray';
      const planned = payload.planned || (payload.source === 'planned' ? payload.data : null);
      const data = payload.data || {};
      const key = planned ? String(planned.date || '') : (data.start_date_local ? dateKeyFromDate(new Date(data.start_date_local)) : todayKey());
      const completed = payload.source === 'strava'
        ? data
        : (planned ? completedFromPlanned(planned) : null);
      const cls = planned
        ? complianceStatus(planned, completed, key).cls
        : (payload.source === 'strava' ? 'unplanned' : 'workout');
      if (cls === 'paired-green') return 'status-green';
      if (cls === 'paired-yellow') return 'status-yellow';
      if (cls === 'paired-orange' || cls === 'paired-red') return 'status-red';
      return 'status-gray';
    }

    function applyModalHeaderTint(payload) {
      const top = document.querySelector('#workoutViewModal .wv-top');
      if (!top) return;
      top.classList.remove('status-green', 'status-yellow', 'status-red', 'status-gray');
      top.classList.add(modalStatusClass(payload));
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

    function buildCtlSeriesForDays(days) {
      const observed = buildObservedDailyTssMap();
      const today = parseDateKey(todayKey());
      const warmup = 120;
      const totalDays = days + warmup;
      const start = new Date(today);
      start.setDate(today.getDate() - totalDays + 1);
      const values = [];
      const dateKeys = [];
      for (let i = 0; i < totalDays; i++) {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        const key = dateKeyFromDate(d);
        values.push(Number(observed[key] || 0));
        dateKeys.push(key);
      }
      let ctlPrev = 0;
      let atlPrev = 0;
      const ctlSeries = [];
      const keySeries = [];
      for (let i = 0; i < values.length; i++) {
        const tss = values[i];
        const ctl = ctlPrev + (tss - ctlPrev) / 42;
        const atl = atlPrev + (tss - atlPrev) / 7;
        ctlPrev = ctl;
        atlPrev = atl;
        if (i >= warmup) {
          ctlSeries.push(ctl);
          keySeries.push(dateKeys[i]);
        }
      }
      return { ctlSeries, keySeries };
    }

    function formatDateShort(key) {
      if (!key) return '';
      const d = parseDateKey(key);
      return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
    }

    function renderFitnessGraph(elId, days) {
      const el = document.getElementById(elId);
      if (!el) return;
      const { ctlSeries, keySeries } = buildCtlSeriesForDays(days);
      if (!ctlSeries.length) { el.innerHTML = '<p class="meta" style="font-size:10px;text-align:center;padding:20px 0;">No data</p>'; return; }
      const W = 200, H = 56;
      const minVal = Math.min(...ctlSeries);
      const maxVal = Math.max(...ctlSeries, minVal + 1);
      const range = maxVal - minVal || 1;
      const pts = ctlSeries.map((v, i) => ({
        x: ctlSeries.length > 1 ? (i / (ctlSeries.length - 1)) * W : W / 2,
        y: H - ((v - minVal) / range) * (H - 8) - 4,
        key: keySeries[i],
        ctl: Math.round(v),
      }));
      const pathD = pts.map((p, i) => `${i === 0 ? 'M' : 'L'}${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
      const fillD = `${pathD} L${W},${H} L0,${H} Z`;
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
      svg.setAttribute('preserveAspectRatio', 'none');
      svg.setAttribute('width', '100%');
      svg.setAttribute('height', String(H));
      svg.style.cssText = 'display:block;width:100%;height:56px;overflow:hidden;';
      const fillPath = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      fillPath.setAttribute('d', fillD);
      fillPath.setAttribute('fill', 'rgba(30,88,209,0.18)');
      svg.appendChild(fillPath);
      const linePath = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      linePath.setAttribute('d', pathD);
      linePath.setAttribute('fill', 'none');
      linePath.setAttribute('stroke', '#1e58d1');
      linePath.setAttribute('stroke-width', '1.8');
      linePath.setAttribute('vector-effect', 'non-scaling-stroke');
      svg.appendChild(linePath);
      const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      dot.setAttribute('r', '4');
      dot.setAttribute('fill', '#1e58d1');
      dot.setAttribute('stroke', '#fff');
      dot.setAttribute('stroke-width', '2');
      dot.style.display = 'none';
      svg.appendChild(dot);
      el.innerHTML = '';
      el.appendChild(svg);
      const tooltip = document.createElement('div');
      tooltip.className = 'fitness-hover-tooltip';
      el.appendChild(tooltip);
      svg.style.cursor = 'crosshair';
      svg.addEventListener('mousemove', (ev) => {
        const rect = svg.getBoundingClientRect();
        const ratio = Math.max(0, Math.min(1, (ev.clientX - rect.left) / rect.width));
        const idx = Math.round(ratio * (pts.length - 1));
        const p = pts[Math.max(0, Math.min(pts.length - 1, idx))];
        if (!p) return;
        dot.setAttribute('cx', String(p.x));
        dot.setAttribute('cy', String(p.y));
        dot.style.display = '';
        const tx = Math.min((p.x / W) * rect.width, rect.width - 110);
        const ty = Math.max((p.y / H) * rect.height - 48, 0);
        tooltip.style.left = `${tx}px`;
        tooltip.style.top = `${ty}px`;
        tooltip.style.display = 'block';
        tooltip.innerHTML = `<span style="opacity:.75">${formatDateShort(p.key)}</span><br>Fitness: <strong>${p.ctl}</strong>`;
      });
      svg.addEventListener('mouseleave', () => { dot.style.display = 'none'; tooltip.style.display = 'none'; });
    }

    function renderFitnessGraphs() {
      renderFitnessGraph('fitnessGraph7', 7);
      renderFitnessGraph('fitnessGraph20', 20);
      renderFitnessGraph('fitnessGraph90', 90);
      renderFitnessGraph('fitnessGraph365', 365);
    }

    function renderPerformanceMetrics() {
      const metrics = buildMetricsToDate(todayKey());
      const ctlEl = document.getElementById('ctlVal');
      const atlEl = document.getElementById('atlVal');
      const tsbEl = document.getElementById('tsbVal');
      if (ctlEl) ctlEl.textContent = String(metrics.ctl);
      if (atlEl) atlEl.textContent = String(metrics.atl);
      if (tsbEl) tsbEl.textContent = metrics.tsb > 0 ? `+${metrics.tsb}` : String(metrics.tsb);
      renderFitnessGraphs();
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

    function renderTopFeelRpe(feel, rpe) {
      const node = document.getElementById('wvTopFeelRpe');
      const sep = document.getElementById('wvHeadStatusSep');
      if (!node) return;
      const f = Number(feel || 0);
      const r = Number(rpe || 0);
      const icon = feelEmoji(f);
      if (icon || r > 0) {
        node.textContent = [icon, r > 0 ? String(r) : ''].filter(Boolean).join(' ');
        node.classList.remove('hidden');
        if (sep) sep.classList.remove('hidden');
      } else {
        node.textContent = '';
        node.classList.add('hidden');
        if (sep) sep.classList.add('hidden');
      }
      renderTopComments();
    }

    function renderTopComments() {
      const node = document.getElementById('wvTopComments');
      if (!node) return;
      const count = (modalDraft && Array.isArray(modalDraft.commentsFeed)) ? modalDraft.commentsFeed.length : 0;
      node.textContent = count > 0 ? `💬 x${count}` : '💬';
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
        renderTopComments();
        return;
      }
      const list = modalDraft.commentsFeed || [];
      if (!list.length) {
        wrap.innerHTML = '';
        renderTopComments();
        return;
      }
      wrap.innerHTML = list.map((c, i) => `
        <div class="comment-item" data-index="${i}">
          <span class="comment-text" title="Double-click to edit">${c.replace(/</g, '&lt;')}</span>
          <button class="comment-delete" type="button" aria-label="Delete comment">🗑</button>
        </div>
      `).join('');
      renderTopComments();
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
      document.getElementById('dEventType').value = existingItem ? (existingItem.event_type || 'Road Running') : 'Road Running';
      document.getElementById('dAvailability').value = existingItem ? (existingItem.availability || 'Unavailable') : 'Unavailable';
      const activePriority = existingItem ? (existingItem.priority || 'C') : 'C';
      document.querySelectorAll('#dEventPriority .seg-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.priority === activePriority);
      });

      document.getElementById('workoutFields').classList.toggle('hidden', selectedKind !== 'workout');
      document.getElementById('eventFields').classList.toggle('hidden', selectedKind !== 'event');
      document.getElementById('availabilityFields').classList.toggle('hidden', selectedKind !== 'availability');
      const isWorkout = selectedKind === 'workout';
      document.getElementById('detailMetricsChips').style.display = isWorkout ? 'contents' : 'none';
      document.querySelector('.detail-right').style.display = isWorkout ? 'block' : 'none';
      document.querySelector('.detail-body').style.gridTemplateColumns = isWorkout ? '1fr 340px' : '1fr';
      document.getElementById('nonWorkoutDescription').style.display = isWorkout ? 'none' : 'block';

      document.getElementById('detailModal').classList.add('open');
      captureDetailInitialState();
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
        const activeP = document.querySelector('#dEventPriority .seg-btn.active');
        payload.priority = activeP ? activeP.dataset.priority : 'C';
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

      await loadData();
      if (closeAfter) {
        closeDetailModal();
      }
    }

    function deleteCurrentDetail() {
      if (!editingItemId) return;
      confirmDelete({ onConfirm: async () => {
        const resp = await fetch(`/calendar-items/${editingItemId}`, { method: 'DELETE' });
        if (!resp.ok) return;
        closeDetailModal();
        await loadData();
      } });
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
        map[key].durationMin += completedDurationMinValue(a);
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
            await loadData();
          },
        });
        opts.push({
          label: 'Delete',
          onClick: () => {
            closeContextMenu();
            confirmDelete({ onConfirm: async () => {
              await fetch(`/calendar-items/${payload.data.id}`, { method: 'DELETE' });
              await loadData();
            } });
          },
        });
      }
      if (payload.source === 'strava') {
        opts.push({
          label: 'Delete',
          onClick: () => {
            closeContextMenu();
            confirmDelete({ onConfirm: async () => {
              await fetch(`/activities/${payload.data.id}`, { method: 'DELETE' });
              await loadData();
            } });
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
            await loadData();
          },
        });
      }

      if (!opts.length) return;
      openContextMenu(ev.clientX, ev.clientY, opts);
    }

    function confirmAndPair(plannedId, stravaId) {
      if (localStorage.getItem('pair_skip_confirm') === 'true') {
        pairWorkouts(plannedId, stravaId);
        return;
      }
      const planned = calendarItems.find(i => String(i.id) === String(plannedId));
      const completed = activities.find(a => String(a.id) === String(stravaId));

      const cDur  = hms(Number(completed && completed.moving_time || 0));
      const cDist = completed ? fmtDistanceMeters(Number(completed.distance || 0)) : '--';
      const cTss  = completed ? Math.round(activityToTss(completed)) : 0;
      const pDur  = planned ? formatDurationClockCompact(Number(planned.duration_min || 0)) : '--';
      const sportKey = planned ? workoutTypeSportKey(planned.workout_type) : 'other';
      const iconSrc = ICON_ASSETS[sportKey] || ICON_ASSETS.other;
      const type = (planned && planned.workout_type) || (completed && completed.type) || 'Workout';

      const overlay = document.createElement('div');
      overlay.className = 'pair-confirm-overlay';
      overlay.innerHTML = `
        <div class="pair-confirm-modal">
          <h3 class="pair-confirm-title">Are you sure you want to pair these workouts?</h3>
          <p class="pair-confirm-sub">Pairing will attach the completed workout to the planned workout. All of your data, descriptions, and comments will remain intact.</p>
          <div class="pair-confirm-preview">
            <div class="pair-preview-col">
              <p class="pair-preview-label">Completed</p>
              <div class="pair-preview-card pair-preview-done">
                <img src="${iconSrc}" class="pair-preview-icon" />
                <strong>${type}</strong>
                <span>${cDur}&#10003;</span>
                <span>${cDist}</span>
                <span>${cTss} TSS</span>
              </div>
              <p class="pair-preview-label" style="margin-top:10px;">Planned</p>
              <div class="pair-preview-card pair-preview-planned">
                <img src="${iconSrc}" class="pair-preview-icon" />
                <span>${pDur}</span>
              </div>
            </div>
            <div class="pair-confirm-arrow">&#8594;</div>
            <div class="pair-preview-col">
              <div class="pair-preview-card pair-preview-merged">
                <img src="${iconSrc}" class="pair-preview-icon" />
                <span>${cDur}&#10003;</span>
                <span>${cDist}</span>
                <span>${cTss} TSS</span>
                <span class="pair-preview-planned-line">P: ${pDur}</span>
              </div>
              <p class="pair-preview-label pair-label-merged">Planned and<br>Completed</p>
            </div>
          </div>
          <label class="pair-confirm-skip">
            <input type="checkbox" id="pairSkipCheck" /> Don't show this again
          </label>
          <div class="pair-confirm-btns">
            <button class="btn-secondary" id="pairCancelBtn">Cancel</button>
            <button class="btn-primary" id="pairConfirmBtn">Pair</button>
          </div>
        </div>`;

      document.body.appendChild(overlay);

      overlay.querySelector('#pairCancelBtn').addEventListener('click', () => overlay.remove());
      overlay.addEventListener('click', (ev) => { if (ev.target === overlay) overlay.remove(); });
      overlay.querySelector('#pairConfirmBtn').addEventListener('click', async () => {
        if (overlay.querySelector('#pairSkipCheck').checked) {
          localStorage.setItem('pair_skip_confirm', 'true');
        }
        overlay.remove();
        await pairWorkouts(plannedId, stravaId);
      });
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
      await loadData();
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

    function workoutDraftSnapshot() {
      const val = (id) => {
        const node = document.getElementById(id);
        return node ? String(node.value || '').trim() : '';
      };
      return {
        title: val('wvTitle'),
        description: val('wvDescription'),
        sportType: modalDraft ? String(modalDraft.sportType || '') : '',
        feel: Number(currentFeel || 0),
        rpeTouched: !!(modalDraft && modalDraft.rpeTouched),
        rpeValue: val('wvRpe'),
        distanceUnit: val('pcDistanceUnit'),
        elevationUnit: val('pcElevationUnit'),
        commentsFeed: modalDraft ? JSON.stringify(modalDraft.commentsFeed || []) : '[]',
        pendingDeleteFit: !!(modalDraft && modalDraft.pendingDeleteFit),
        uploadedNow: !!(modalDraft && modalDraft.uploadedNow),
        fields: {
          pcDurPlan: val('pcDurPlan'),
          pcDurComp: val('pcDurComp'),
          pcDistPlan: val('pcDistPlan'),
          pcDistComp: val('pcDistComp'),
          pcAvgSpeedPlan: val('pcAvgSpeedPlan'),
          pcAvgSpeedComp: val('pcAvgSpeedComp'),
          pcCaloriesPlan: val('pcCaloriesPlan'),
          pcCaloriesComp: val('pcCaloriesComp'),
          pcElevPlan: val('pcElevPlan'),
          pcElevComp: val('pcElevComp'),
          pcTssPlan: val('pcTssPlan'),
          pcTssComp: val('pcTssComp'),
          pcIfPlan: val('pcIfPlan'),
          pcIfComp: val('pcIfComp'),
          pcNpComp: val('pcNpComp'),
          pcWorkPlan: val('pcWorkPlan'),
          pcWorkComp: val('pcWorkComp'),
          wvHrMin: val('wvHrMin'),
          wvHrAvg: val('wvHrAvg'),
          wvHrMax: val('wvHrMax'),
          wvPowerMin: val('wvPowerMin'),
          wvPowerAvg: val('wvPowerAvg'),
          wvPowerMax: val('wvPowerMax'),
        },
      };
    }

    function hasUnsavedWorkoutChanges() {
      if (!window.currentWorkoutPayload || !modalDraft) return false;
      if (analyzeState && analyzeState.pendingDirty) return true;
      const start = modalDraft.initialSnapshot;
      if (!start) return false;
      const now = workoutDraftSnapshot();
      return JSON.stringify(start) !== JSON.stringify(now);
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
          await loadData();
        };
      }
      const delBtn = document.getElementById('wvDeleteFitBtn');
      if (delBtn) {
        delBtn.onclick = () => {
          if (!modalDraft) return;
          confirmDelete({ onConfirm: () => {
            modalDraft.pendingDeleteFit = true;
            renderWorkoutFiles(window.currentWorkoutPayload);
          } });
        };
      }
    }

    function openWorkoutModal(payload) {
      const modalSession = ++workoutModalSession;
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
          if (window.currentWorkoutPayload) {
            if (window.currentWorkoutPayload.planned) {
              window.currentWorkoutPayload.planned.workout_type = modalDraft.sportType;
            } else if (window.currentWorkoutPayload.data) {
              window.currentWorkoutPayload.data.workout_type = modalDraft.sportType;
              window.currentWorkoutPayload.data.type = modalDraft.sportType;
            }
            renderWorkoutSummary(window.currentWorkoutPayload);
          }
          sportMenu.classList.add('hidden');
        });
      });
      if (hasFile) {
        renderWorkoutAnalyze(payload, modalSession);
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
      applyModalHeaderTint(payload);
      modalDraft.initialSnapshot = workoutDraftSnapshot();
      document.getElementById('workoutViewModal').classList.add('open');
    }

    async function closeWorkoutModal(discard = true) {
      workoutModalSession += 1;
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
      analyzeState = null;
      window.currentWorkoutPayload = null;
      fitUploadContext = 'global';
      fitUploadTargetActivityId = null;
    }

    async function persistWorkoutView() {
      const payload = window.currentWorkoutPayload;
      if (!payload) return { saved: false, emptyDraft: false };
      if (analyzeState && analyzeState.pendingDirty) {
        setWorkoutMode('analyze');
        const applyNow = await confirmApplyChanges();
        if (applyNow && typeof analyzeState.applyPending === 'function') {
          await analyzeState.applyPending();
        } else if (typeof analyzeState.cancelPending === 'function') {
          analyzeState.cancelPending();
        }
      }
      const data = payload.data || {};
      const targetPlanned = payload.planned || (payload.source === 'planned' ? data : null);
      recalcIfTssRows();
      const description = document.getElementById('wvDescription').value;
      const commentsFeed = modalDraft ? modalDraft.commentsFeed.slice() : [];
      const comments = commentsFeed.length ? commentsFeed[commentsFeed.length - 1] : '';
      const sport = (modalDraft && modalDraft.sportType) || 'Other';
      const distanceUnitLocal = document.getElementById('pcDistanceUnit').value || 'km';
      const elevationUnitLocal = document.getElementById('pcElevationUnit').value || 'm';
      const persistedDistanceUnit = distanceUnitLocal;
      const persistedElevationUnit = elevationUnitLocal;
      const plannedDuration = parseDurationToMin(document.getElementById('pcDurPlan').value);
      const plannedDistanceM = fromDisplayDistanceToMeters(document.getElementById('pcDistPlan').value, distanceUnitLocal);
      const plannedElevationM = fromDisplayElevationToMeters(document.getElementById('pcElevPlan').value, elevationUnitLocal);
      const plannedTss = Number(document.getElementById('pcTssPlan').value || 0);
      const plannedIf = Number(document.getElementById('pcIfPlan').value || 0);
      const plannedAvgSpeedDisplay = Number(document.getElementById('pcAvgSpeedPlan').value || 0);
      const plannedAvgSpeed = Number.isFinite(plannedAvgSpeedDisplay) && plannedAvgSpeedDisplay > 0
        ? (distanceUnitLocal === 'mi'
          ? (plannedAvgSpeedDisplay / 2.23694)
          : (distanceUnitLocal === 'km'
            ? (plannedAvgSpeedDisplay / 3.6)
            : plannedAvgSpeedDisplay))
        : 0;
      const plannedCalories = Number(document.getElementById('pcCaloriesPlan').value || 0);
      const plannedWorkKj = Number(document.getElementById('pcWorkPlan').value || 0);
      const completedDuration = parseDurationToMin(document.getElementById('pcDurComp').value);
      const completedDistanceM = fromDisplayDistanceToMeters(document.getElementById('pcDistComp').value, distanceUnitLocal);
      const completedElevationM = fromDisplayElevationToMeters(document.getElementById('pcElevComp').value, elevationUnitLocal);
      const completedTss = Number(document.getElementById('pcTssComp').value || 0);
      const completedIf = Number(document.getElementById('pcIfComp').value || 0);
      const completedNp = Number(document.getElementById('pcNpComp').value || 0);
      const completedWorkKj = Number(document.getElementById('pcWorkComp').value || 0);
      const completedCalories = Number(document.getElementById('pcCaloriesComp').value || 0);
      const completedHrMin = Number(document.getElementById('wvHrMin').value || 0);
      const completedHrAvg = Number(document.getElementById('wvHrAvg').value || 0);
      const completedHrMax = Number(document.getElementById('wvHrMax').value || 0);
      const completedPowerMin = Number(document.getElementById('wvPowerMin').value || 0);
      const completedPowerAvg = Number(document.getElementById('wvPowerAvg').value || 0);
      const completedPowerMax = Number(document.getElementById('wvPowerMax').value || 0);
      const completedAvgSpeedDisplay = Number(document.getElementById('pcAvgSpeedComp').value || 0);
      const completedAvgSpeed = Number.isFinite(completedAvgSpeedDisplay) && completedAvgSpeedDisplay > 0
        ? (distanceUnitLocal === 'mi'
          ? (completedAvgSpeedDisplay / 2.23694)
          : (distanceUnitLocal === 'km'
            ? (completedAvgSpeedDisplay / 3.6)
            : completedAvgSpeedDisplay))
        : 0;
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
        return { saved: false, emptyDraft: true };
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
            distance_unit: persistedDistanceUnit,
            elevation_unit: persistedElevationUnit,
            planned_tss: plannedTss,
            planned_if: plannedIf,
            planned_avg_speed: plannedAvgSpeed,
            planned_calories: plannedCalories,
            planned_work_kj: plannedWorkKj,
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
            completed_np: completedNp,
            completed_work_kj: completedWorkKj,
            completed_calories: completedCalories,
            completed_avg_speed: completedAvgSpeed,
            completed_hr_min: completedHrMin,
            completed_hr_avg: completedHrAvg,
            completed_hr_max: completedHrMax,
            completed_power_min: completedPowerMin,
            completed_power_avg: completedPowerAvg,
            completed_power_max: completedPowerMax,
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
            distance_unit: persistedDistanceUnit,
            elevation_unit: persistedElevationUnit,
            planned_tss: plannedTss,
            planned_if: plannedIf,
            planned_avg_speed: plannedAvgSpeed,
            planned_calories: plannedCalories,
            planned_work_kj: plannedWorkKj,
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
            completed_np: completedNp,
            completed_work_kj: completedWorkKj,
            completed_calories: completedCalories,
            completed_avg_speed: completedAvgSpeed,
            completed_hr_min: completedHrMin,
            completed_hr_avg: completedHrAvg,
            completed_hr_max: completedHrMax,
            completed_power_min: completedPowerMin,
            completed_power_avg: completedPowerAvg,
            completed_power_max: completedPowerMax,
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
            tss_source: (modalDraft && modalDraft.tssSource) || '',
            duration_min: plannedDuration,
            distance_km: plannedDistanceM / 1000,
            distance_m: plannedDistanceM,
            elevation_m: plannedElevationM,
            distance_unit: persistedDistanceUnit,
            elevation_unit: persistedElevationUnit,
            planned_tss: plannedTss,
            planned_if: plannedIf,
            planned_avg_speed: plannedAvgSpeed,
            planned_calories: plannedCalories,
            planned_work_kj: plannedWorkKj,
            title: document.getElementById('wvTitle').value.trim(),
            type: sport,
          }),
        });
      }
      return { saved: true, emptyDraft: false };
    }

    async function handleSave() {
      const result = await persistWorkoutView();
      if (result && result.saved && modalDraft) {
        modalDraft.pendingDeleteFit = false;
        modalDraft.uploadedNow = false;
        modalDraft.initialSnapshot = workoutDraftSnapshot();
      }
    }

    async function handleSaveAndClose() {
      const result = await persistWorkoutView();
      await closeWorkoutModal(false);
      if (result && (result.saved || result.emptyDraft)) {
        await loadData();
      }
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
      applyModalHeaderTint(payload);
      const data = payload.data || {};
      const parentPlanned = payload.planned || null;
      const explicitCompleted = parentPlanned ? completedFromPlanned(parentPlanned) : completedFromPlanned(data);
      const completedDurationMin = payload.source === 'strava'
        ? completedDurationMinValue(data)
        : explicitCompleted ? Number(explicitCompleted.moving_time || 0) / 60 : 0;
      const completedDistanceM = payload.source === 'strava'
        ? Number(data.distance || 0)
        : explicitCompleted ? Number(explicitCompleted.distance || 0) : Number((parentPlanned || data).completed_distance_m || 0);
      const completedTssRaw = payload.source === 'strava'
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
      const isCycling = String(typeLabel || '').toLowerCase().includes('bike') || String(typeLabel || '').toLowerCase().includes('ride');
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
      const savedFeel = (parentPlanned && parentPlanned.feel) || data.feel || 0;
      setFeelValue(savedFeel);
      renderTopFeelRpe(savedFeel, savedRpe);

      const plannedDuration = parentPlanned ? Number(parentPlanned.duration_min || 0) : Number(data.duration_min || 0);
      const plannedDistanceM = Number((plannedObj.distance_m || 0) || (Number(plannedObj.distance_km || 0) * 1000));
      const plannedElevationM = Number(plannedObj.elevation_m || 0);
      const completedElevationM = payload.source === 'strava'
        ? Number(data.elev_gain_m || 0)
        : Number(plannedObj.completed_elevation_m || 0);
      const plannedTss = Number(plannedObj.planned_tss || 0) || (parentPlanned ? itemToTss(parentPlanned) : itemToTss(data));
      const plannedIf = plannedIF(plannedObj);
      const completedIfRaw = payload.source === 'strava' ? activityIF(data) : completedIF({
        completed_if: plannedObj.completed_if,
        completed_tss: completedTssRaw,
        moving_time: completedDurationMin * 60,
        avg_power: data.avg_power,
        type: plannedObj.workout_type || data.type,
      });
      const fitComputedNp = Number((data.np_value ?? data.normalized_power ?? data.np ?? plannedObj.completed_np) || 0);
      const fitComputedIf = Number((data.if_value ?? completedIfRaw) || 0);
      const fitComputedTss = Number((data.tss_override ?? completedTssRaw) || 0);
      const completedNp = fitComputedNp > 0 ? fitComputedNp : null;
      let completedIf = fitComputedIf > 0 ? fitComputedIf : null;
      const powerTssValue = (isCycling && !completedNp) ? null : (fitComputedTss > 0 ? fitComputedTss : null);
      const hrTssValue = calcHrTss(data);
      const hasBothTss = !!(powerTssValue && hrTssValue);
      let activeTssSource = data.tss_source || (powerTssValue ? 'power' : 'hr');
      let completedTss = activeTssSource === 'hr' ? (hrTssValue || powerTssValue) : (powerTssValue || hrTssValue);
      if (!completedTss) completedTss = null;
      if (isCycling && !completedNp) completedIf = null;

      // TSS source toggle — shown only when both power and hrTSS are available
      const tssSourceRow = document.getElementById('wvTssSourceRow');
      tssSourceRow.style.display = hasBothTss ? '' : 'none';
      if (hasBothTss) {
        document.querySelectorAll('#wvTssSourceToggle .seg-btn').forEach(btn => {
          btn.classList.toggle('active', btn.dataset.source === activeTssSource);
          btn.onclick = () => {
            activeTssSource = btn.dataset.source;
            completedTss = activeTssSource === 'hr' ? hrTssValue : powerTssValue;
            document.querySelectorAll('#wvTssSourceToggle .seg-btn').forEach(b => b.classList.toggle('active', b.dataset.source === activeTssSource));
            document.getElementById('pcTssComp').value = completedTss ? String(Math.round(completedTss)) : '';
            document.getElementById('wvHeaderTss').textContent = completedTss ? `${Math.round(completedTss)} TSS` : '-- TSS';
            if (modalDraft) modalDraft.tssSource = activeTssSource;
          };
        });
        if (modalDraft) modalDraft.tssSource = activeTssSource;
      }
      const completedWorkKj = Number((data.work_kj ?? plannedObj.completed_work_kj ?? 0) || 0);
      const completedCalories = Number((data.calories ?? plannedObj.completed_calories ?? 0) || 0);
      const plannedAvgSpeed = Number(plannedObj.planned_avg_speed || 0);
      const plannedCalories = Number(plannedObj.planned_calories || 0);
      const plannedWorkKj = Number(plannedObj.planned_work_kj || 0);
      const compAvgSpeed = payload.source === 'strava'
        ? Number(data.avg_speed || 0)
        : Number(plannedObj.completed_avg_speed || 0);
      const speedLabel = distanceUnitLocal === 'mi' ? 'mph' : (distanceUnitLocal === 'm' ? 'm/s' : 'km/h');
      const speedToDisplay = (v) => {
        if (!Number.isFinite(v) || v <= 0) return '';
        if (distanceUnitLocal === 'mi') return (v * 2.23694).toFixed(1);
        if (distanceUnitLocal === 'm') return v.toFixed(2);
        return (v * 3.6).toFixed(1);
      };

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
      document.getElementById('pcAvgSpeedPlan').value = speedToDisplay(plannedAvgSpeed);
      document.getElementById('pcAvgSpeedComp').value = speedToDisplay(compAvgSpeed);
      document.getElementById('pcCaloriesPlan').value = plannedCalories > 0 ? String(Math.round(plannedCalories)) : '';
      document.getElementById('pcCaloriesComp').value = completedCalories > 0 ? String(Math.round(completedCalories)) : '';
      document.getElementById('pcNpPlan').value = '';
      document.getElementById('pcNpComp').value = completedNp ? String(Math.round(completedNp)) : '';
      document.getElementById('pcWorkPlan').value = plannedWorkKj > 0 ? String(Math.round(plannedWorkKj)) : '';
      document.getElementById('pcWorkComp').value = completedWorkKj > 0 ? String(Math.round(completedWorkKj)) : '';
      document.querySelector('#pcAvgSpeedPlan').closest('.tp-row').querySelector('.tp-unit').textContent = speedLabel;
      document.getElementById('wvHrMin').value = (data.min_hr ?? plannedObj.completed_hr_min) ? String(Math.round(data.min_hr ?? plannedObj.completed_hr_min)) : '';
      document.getElementById('wvHrAvg').value = (data.avg_hr ?? plannedObj.completed_hr_avg) ? String(Math.round(data.avg_hr ?? plannedObj.completed_hr_avg)) : '';
      document.getElementById('wvHrMax').value = (data.max_hr ?? plannedObj.completed_hr_max) ? String(Math.round(data.max_hr ?? plannedObj.completed_hr_max)) : '';
      document.getElementById('wvPowerMin').value = (data.min_power ?? plannedObj.completed_power_min) ? String(Math.round(data.min_power ?? plannedObj.completed_power_min)) : '';
      document.getElementById('wvPowerAvg').value = (data.avg_power ?? plannedObj.completed_power_avg) ? String(Math.round(data.avg_power ?? plannedObj.completed_power_avg)) : '';
      document.getElementById('wvPowerMax').value = (data.max_power ?? plannedObj.completed_power_max) ? String(Math.round(data.max_power ?? plannedObj.completed_power_max)) : '';
      const hrRow = document.getElementById('wvHrMin').closest('.tp-minmax-row');
      if (hrRow) hrRow.style.gridTemplateColumns = '110px 1.5fr 1.5fr 1.5fr 80px';
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
        el.disabled = false;
        el.readOnly = false;
        el.classList.remove('no-entry');
        el.oninput = null;
      });
      const completedFieldIds = ['pcDurComp', 'pcDistComp', 'pcElevComp', 'pcTssComp', 'pcIfComp', 'pcAvgSpeedComp', 'pcCaloriesComp', 'pcNpComp', 'pcWorkComp', 'wvHrMin', 'wvHrAvg', 'wvHrMax', 'wvPowerAvg', 'wvPowerMax'];
      const futureWorkout = isFutureDateKey(parentPlanned ? parentPlanned.date : data.date);
      completedFieldIds.forEach((id) => {
        const el = document.getElementById(id);
        if (!el) return;
        el.disabled = futureWorkout;
        el.classList.toggle('muted', futureWorkout);
        el.classList.toggle('no-entry', futureWorkout);
      });
      const powerMinNode = document.getElementById('wvPowerMin');
      powerMinNode.disabled = true;
      powerMinNode.readOnly = true;
      powerMinNode.classList.add('muted', 'no-entry');
      const npPlanNode = document.getElementById('pcNpPlan');
      npPlanNode.disabled = true;
      npPlanNode.readOnly = true;
      npPlanNode.classList.add('muted', 'no-entry');

      const cyclingRows = ['pcAvgSpeedPlan', 'pcCaloriesPlan', 'pcNpPlan', 'pcWorkPlan'];
      cyclingRows.forEach((id) => {
        const row = document.getElementById(id).closest('.tp-row');
        if (row) row.style.display = isCycling ? '' : 'none';
      });

      function recalcDerivedLive() {
        const pDurMin = parseDurationToMin(document.getElementById('pcDurPlan').value);
        const cDurMin = parseDurationToMin(document.getElementById('pcDurComp').value);
        const pDistM = fromDisplayDistanceToMeters(document.getElementById('pcDistPlan').value, distanceUnitLocal);
        const cDistM = fromDisplayDistanceToMeters(document.getElementById('pcDistComp').value, distanceUnitLocal);
        if (pDurMin > 0 && pDistM > 0) {
          const pSpeedMps = pDistM / (pDurMin * 60);
          document.getElementById('pcAvgSpeedPlan').value = speedToDisplay(pSpeedMps);
        }
        if (cDurMin > 0 && cDistM > 0) {
          const cSpeedMps = cDistM / (cDurMin * 60);
          document.getElementById('pcAvgSpeedComp').value = speedToDisplay(cSpeedMps);
        }
        if (!data.fit_id) {
          const ftpRide = Number((appSettings.ftp || {}).ride || 0);
          const avgPower = Number(document.getElementById('wvPowerAvg').value || 0);
          if (avgPower > 0 && cDurMin > 0) {
            const durS = cDurMin * 60;
            const np = avgPower;
            document.getElementById('pcNpComp').value = String(Math.round(np));
            document.getElementById('pcWorkComp').value = String(Math.round((np * durS) / 1000));
            if (!document.getElementById('pcCaloriesComp').value) {
              document.getElementById('pcCaloriesComp').value = String(Math.round((np * durS) / 1000));
            }
            if (ftpRide > 0) {
              const ifv = np / ftpRide;
              const tss = (durS * np * ifv) / (ftpRide * 3600) * 100;
              document.getElementById('pcIfComp').value = ifv.toFixed(2);
              document.getElementById('pcTssComp').value = String(Math.round(tss));
            }
          }
        }
      }

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
      ['pcDurPlan', 'pcDurComp', 'pcDistPlan', 'pcDistComp', 'wvPowerAvg'].forEach((id) => {
        const n = document.getElementById(id);
        if (!n) return;
        n.oninput = () => {
          recalcIfTssRows();
          recalcDerivedLive();
        };
      });
      recalcDerivedLive();
    }

    async function renderWorkoutAnalyze(payload, modalSession = workoutModalSession) {
      if (modalSession !== workoutModalSession) return;
      const data = payload.data || {};
      const chart = document.getElementById('wvChart');
      const lapBody = document.querySelector('#wvLapTable tbody');
      const statsNode = document.getElementById('wvSelectionKv');
      if (!data.fit_id) {
        statsNode.innerHTML = '<div>No FIT stream for this workout.</div>';
        lapBody.innerHTML = '';
        chart.innerHTML = '';
        return;
      }

      const resp = await fetch(`/fit/${data.fit_id}`);
      if (!resp.ok) {
        statsNode.innerHTML = '<div>Could not load FIT data.</div>';
        lapBody.innerHTML = '';
        chart.innerHTML = '';
        return;
      }
      const fit = await resp.json();
      if (modalSession !== workoutModalSession) return;
      const series = Array.isArray(fit.series) ? fit.series : [];
      const laps = Array.isArray(fit.laps) ? fit.laps : [];
      const summary = fit.summary || {};
      if (!series.length) {
        statsNode.innerHTML = '<div>No FIT points available.</div>';
        lapBody.innerHTML = '';
        chart.innerHTML = '';
        return;
      }

      const baseMs = new Date(series[0].timestamp).getTime();
      const pts = series.map((p) => ({
        t: timeToSec(p.timestamp, baseMs),
        timestamp: p.timestamp,
        heart_rate: num(p.heart_rate),
        speed: num(p.speed),
        distance: num(p.distance),
        cadence: num(p.cadence),
        power: num(p.power),
      }));
      const totalSec = Math.max(1, pts[pts.length - 1].t - pts[0].t);

      const lineMeta = [
        { key: 'cadence', color: '#f39b1f', label: 'RPM', side: 'left', unit: 'RPM' },
        { key: 'heart_rate', color: '#f35353', label: 'BPM', side: 'right', unit: 'BPM' },
        { key: 'power', color: '#cc44f0', label: 'W', side: 'left', unit: 'W' },
        { key: 'speed', color: '#3fa144', label: 'MPH', side: 'right', unit: (distanceUnit === 'mi' ? 'MPH' : 'KM/H') },
      ];

      analyzeState = {
        pts,
        laps,
        totalSec,
        wStart: 0,
        wEnd: totalSec,
        selectionMode: 'none',
        selection: null,
        lapHighlightRanges: [],
        hiddenChannels: new Set(),
        deletedChannels: new Set(),
        cursorSvgX: -1,
        smoothingStep: 2,
        smoothedCache: new Map(),
        zoomStack: [],
        pendingDirty: false,
        persistedDeleted: new Set(),
        persistedCuts: [],
        workDeleted: new Set(),
        workCuts: [],
        pendingCuts: [],
        startTrim: 0,
        selectedLapKeys: new Set(),
      };

      const w = 1200;
      const h = 180;
      const left = 72;
      const right = 72;
      const top = 14;
      const bottom = 30;
      const cw = w - left - right;
      const ch = h - top - bottom;

      chart.setAttribute('viewBox', `0 0 ${w} ${h}`);
      chart.setAttribute('height', String(h));

      const zoomBtn = document.getElementById('wvZoomBtn');
      const cutBtn = document.getElementById('wvCutBtn');
      const unzoomBtn = document.getElementById('wvUnzoomBtn');
      const smoothingStep = document.getElementById('wvSmoothingStep');
      const channelNode = document.getElementById('wvAnalyzeChannelsTop');
      const selectionTitle = document.getElementById('wvSelectionTitle');
      let activeChannelPopup = null;
      const applyBar = document.getElementById('wvAnalyzeApplyBar');
      const applyBtn = document.getElementById('wvApplyAnalyzeBtn');
      const cancelBtn = document.getElementById('wvCancelAnalyzeBtn');
      const applyHost = document.getElementById('wvAnalyzeApplyHost');

      const persisted = ((payload.planned && payload.planned.analysis_edits) || data.analysis_edits || {});
      const persistedDeleted = Array.isArray(persisted.deletedChannels) ? persisted.deletedChannels : [];
      const persistedCuts = Array.isArray(persisted.cuts) ? persisted.cuts : [];
      analyzeState.persistedDeleted = new Set(persistedDeleted.map((x) => String(x)));
      analyzeState.workDeleted = new Set(analyzeState.persistedDeleted);
      analyzeState.persistedCuts = persistedCuts
        .map((c) => ({ startSec: Number(c.startSec || 0), endSec: Number(c.endSec || 0) }))
        .filter((c) => c.endSec > c.startSec);
      analyzeState.workCuts = analyzeState.persistedCuts.map((c) => ({ ...c }));

      const existsChannel = (key) => pts.some((p) => p[key] !== null);
      const toDisplaySpeed = (v) => (distanceUnit === 'mi' ? (v * 2.23694) : (v * 3.6));
      const toDisplayPace = (speedMps) => {
        if (!speedMps || speedMps <= 0) return null;
        const sec = (distanceUnit === 'mi' ? 1609.344 : 1000) / speedMps;
        const m = Math.floor(sec / 60);
        const s = Math.round(sec % 60);
        return `${m}:${String(s).padStart(2, '0')}`;
      };
      const speedUnitLabel = distanceUnit === 'mi' ? 'MPH' : 'KM/H';
      const paceUnitLabel = distanceUnit === 'mi' ? 'min/mi' : 'min/km';

      function normalizeCuts(cuts) {
        const inCuts = (cuts || [])
          .map((c) => ({ startSec: Number(c.startSec || 0), endSec: Number(c.endSec || 0) }))
          .filter((c) => c.endSec > c.startSec)
          .sort((a, b) => a.startSec - b.startSec);
        const out = [];
        inCuts.forEach((c) => {
          if (!out.length) { out.push(c); return; }
          const last = out[out.length - 1];
          if (c.startSec <= last.endSec) last.endSec = Math.max(last.endSec, c.endSec);
          else out.push(c);
        });
        return out;
      }

      function inCut(origT, includePending = true) {
        const applied = analyzeState.workCuts.some((c) => origT >= c.startSec && origT <= c.endSec);
        if (applied) return true;
        if (!includePending) return false;
        return analyzeState.pendingCuts.some((c) => origT >= c.startSec && origT <= c.endSec);
      }

      function displayFromOriginal(origT) {
        return Math.max(0, origT - analyzeState.startTrim);
      }

      function originalFromDisplay(displayT) {
        return Math.max(0, displayT) + analyzeState.startTrim;
      }

      function recomputeCutState() {
        analyzeState.workCuts = normalizeCuts(analyzeState.workCuts);
        let start = 0;
        let end = totalSec;
        let changed = true;
        while (changed) {
          changed = false;
          analyzeState.workCuts.forEach((c) => {
            if (c.startSec <= start + 0.01 && c.endSec > start + 0.01) {
              const next = Math.max(start, c.endSec);
              if (next !== start) { start = next; changed = true; }
            }
            if (c.endSec >= end - 0.01 && c.startSec < end - 0.01) {
              const next = Math.min(end, c.startSec);
              if (next !== end) { end = next; changed = true; }
            }
          });
        }
        analyzeState.startTrim = start;
        analyzeState.endTrim = end;
      }

      function effectivePoints() {
        return pts
          .filter((p) => p.t >= analyzeState.startTrim && p.t <= analyzeState.endTrim && !inCut(p.t, false))
          .map((p) => ({ ...p, dT: displayFromOriginal(p.t) }));
      }

      function syncPendingState() {
        analyzeState.pendingCuts = normalizeCuts(analyzeState.pendingCuts);
        const deletedNow = Array.from(analyzeState.workDeleted || []).sort();
        const deletedPersisted = Array.from(analyzeState.persistedDeleted || []).sort();
        const deletedChanged = deletedNow.length !== deletedPersisted.length
          || deletedNow.some((key, i) => key !== deletedPersisted[i]);
        const cutsChanged = analyzeState.pendingCuts.length > 0;
        analyzeState.pendingDirty = deletedChanged || cutsChanged;
        if (applyHost) applyHost.classList.toggle('hidden', !analyzeState.pendingDirty);
        if (applyBar) applyBar.classList.toggle('hidden', !analyzeState.pendingDirty);
      }

      async function persistAnalyzeEdits() {
        const mergedCuts = normalizeCuts([...(analyzeState.workCuts || []), ...(analyzeState.pendingCuts || [])]);
        analyzeState.workCuts = mergedCuts;
        analyzeState.pendingCuts = [];
        recomputeCutState();
        const eff = effectivePoints();
        const lapDurationSec = (lapRows || []).reduce((sum, lap) => {
          const startSec = Math.max(0, timeToSec(lap.start || series[0].timestamp, baseMs));
          const endSec = Math.max(startSec + 1, timeToSec(lap.end || series[series.length - 1].timestamp, baseMs));
          return sum + lapKeptDuration(startSec, endSec);
        }, 0);
        const effectiveDurationSec = Math.max(0, lapDurationSec || (eff.length ? eff[eff.length - 1].dT : 0));
        const effectiveDurationMin = effectiveDurationSec / 60;
        analyzeState.wStart = 0;
        analyzeState.wEnd = Math.max(1, (eff.length ? eff[eff.length - 1].dT : totalSec));
        const edits = {
          deletedChannels: Array.from(analyzeState.workDeleted),
          cuts: mergedCuts.map((c) => ({ startSec: c.startSec, endSec: c.endSec })),
        };
        if (payload.planned && payload.planned.id) {
          const body = { ...payload.planned, analysis_edits: edits, completed_duration_min: effectiveDurationMin };
          await fetch(`/calendar-items/${payload.planned.id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
          });
          payload.planned.analysis_edits = edits;
          payload.planned.completed_duration_min = effectiveDurationMin;
          if (payload.data) payload.data.completed_duration_min = effectiveDurationMin;
        } else if (payload.source === 'strava' && payload.data && payload.data.id) {
          await fetch(`/activities/${payload.data.id}/meta`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ analysis_edits: edits, completed_duration_min: effectiveDurationMin }),
          });
          payload.data.analysis_edits = edits;
          payload.data.completed_duration_min = effectiveDurationMin;
          payload.data.moving_time = effectiveDurationSec;
        }
        analyzeState.persistedDeleted = new Set(edits.deletedChannels);
        analyzeState.persistedCuts = edits.cuts.map((c) => ({ ...c }));
        syncPendingState();
        renderLapRows();
        renderWorkoutSummary(payload);
        if (modalDraft) modalDraft.initialSnapshot = workoutDraftSnapshot();
        renderMain();
      }

      recomputeCutState();
      const initialEffPts = effectivePoints();
      const initialEffEnd = initialEffPts.length ? initialEffPts[initialEffPts.length - 1].dT : totalSec;
      analyzeState.wStart = 0;
      analyzeState.wEnd = Math.max(1, initialEffEnd);

      function fmtTimeShort(iso) {
        if (!iso) return '';
        const dt = new Date(iso);
        if (Number.isNaN(dt.getTime())) return '';
        return dt.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
      }

      function smoothedValues(key) {
        const step = Math.max(1, Number(analyzeState.smoothingStep || 2));
        const cacheKey = `${key}:${step}`;
        if (analyzeState.smoothedCache.has(cacheKey)) return analyzeState.smoothedCache.get(cacheKey);
        if (step <= 1) {
          const raw = pts.map((p) => p[key]);
          analyzeState.smoothedCache.set(cacheKey, raw);
          return raw;
        }
        const half = Math.max(1, (step * step) - 1);
        const out = new Array(pts.length).fill(null);
        for (let i = 0; i < pts.length; i += 1) {
          let sum = 0;
          let count = 0;
          const start = Math.max(0, i - half);
          const end = Math.min(pts.length - 1, i + half);
          for (let j = start; j <= end; j += 1) {
            const v = pts[j][key];
            if (v !== null) {
              sum += v;
              count += 1;
            }
          }
          out[i] = count ? (sum / count) : null;
        }
        analyzeState.smoothedCache.set(cacheKey, out);
        return out;
      }

      function valueForAxis(key, value) {
        if (value == null) return null;
        if (key === 'speed') return toDisplaySpeed(value);
        return value;
      }

      function visibleChannels() {
        return lineMeta.filter((m) => existsChannel(m.key) && !analyzeState.hiddenChannels.has(m.key) && !analyzeState.workDeleted.has(m.key));
      }

      function findNearestPointAt(time) {
        const eff = effectivePoints();
        let nearest = eff[0] || pts[0];
        let best = Infinity;
        for (let i = 0; i < eff.length; i += 1) {
          const d = Math.abs(eff[i].dT - time);
          if (d < best) {
            best = d;
            nearest = eff[i];
          }
        }
        return nearest;
      }

      function timeToX(t) {
        return left + ((t - analyzeState.wStart) / Math.max(1, analyzeState.wEnd - analyzeState.wStart)) * cw;
      }

      function plotMouseToTime(ev) {
        const rect = chart.getBoundingClientRect();
        const x = ev.clientX - rect.left;
        const clamped = Math.max(0, Math.min(rect.width, x));
        const leftPx = (left / w) * rect.width;
        const plotWidthPx = (cw / w) * rect.width;
        const plotX = Math.max(0, Math.min(plotWidthPx, clamped - leftPx));
        const frac = plotX / Math.max(1, plotWidthPx);
        return {
          time: analyzeState.wStart + frac * (analyzeState.wEnd - analyzeState.wStart),
          svgX: left + frac * cw,
        };
      }

      function statsForRange(startT, endT) {
        const rangeStartOrig = originalFromDisplay(Math.min(startT, endT));
        const rangeEndOrig = originalFromDisplay(Math.max(startT, endT));
        const activeCuts = normalizeCuts([...(analyzeState.workCuts || []), ...(analyzeState.pendingCuts || [])]);
        const globalKept = [];
        let cursor = 0;
        activeCuts.forEach((c) => {
          if (c.startSec > cursor) globalKept.push({ startSec: cursor, endSec: c.startSec });
          cursor = Math.max(cursor, c.endSec);
        });
        if (cursor < totalSec) globalKept.push({ startSec: cursor, endSec: totalSec });
        const selectedKept = [];
        globalKept.forEach((seg) => {
          const s = Math.max(seg.startSec, rangeStartOrig);
          const e = Math.min(seg.endSec, rangeEndOrig);
          if (e > s) selectedKept.push({ startSec: s, endSec: e });
        });
        const selectedPts = pts.filter((p) => selectedKept.some((r) => p.t >= r.startSec && p.t <= r.endSec));
        const duration = selectedKept.reduce((sum, r) => sum + Math.max(0, r.endSec - r.startSec), 0);
        const effectiveTotal = globalKept.reduce((sum, r) => sum + Math.max(0, r.endSec - r.startSec), 0) || 1;
        const frac = duration / effectiveTotal;
        const distance = selectedKept.reduce((sum, r) => {
          const seg = pts.filter((p) => p.t >= r.startSec && p.t <= r.endSec && p.distance !== null);
          if (seg.length > 1) {
            return sum + Math.max(0, seg[seg.length - 1].distance - seg[0].distance);
          }
          return sum;
        }, 0);
        const totalTss = Number(summary.tss || data.tss_override || activityToTss(data) || 0);
        const tss = totalTss > 0 ? (totalTss * frac) : null;
        const minVal = (k) => {
          const vals = selectedPts.map((p) => p[k]).filter((v) => v !== null);
          return vals.length ? Math.min(...vals) : null;
        };
        const mean = (k) => {
          const vals = selectedPts.map((p) => p[k]).filter((v) => v !== null);
          return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
        };
        const maxVal = (k) => {
          const vals = selectedPts.map((p) => p[k]).filter((v) => v !== null);
          return vals.length ? Math.max(...vals) : null;
        };
        return { duration, distance, tss, minVal, mean, maxVal };
      }

      function statsForDiscreteRanges(ranges) {
        const validRanges = (Array.isArray(ranges) ? ranges : [])
          .map((r) => ({ startSec: Number(r.startSec || 0), endSec: Number(r.endSec || 0) }))
          .filter((r) => r.endSec > r.startSec);
        if (!validRanges.length) {
          return statsForRange(analyzeState.wStart, analyzeState.wEnd);
        }
        const intervalPoints = (startSec, endSec) => pts.filter((p) => p.t >= startSec && p.t <= endSec && !inCut(p.t));
        const duration = validRanges.reduce((sum, r) => sum + lapKeptDuration(r.startSec, r.endSec), 0);
        const distance = validRanges.reduce((sum, r) => {
          const seg = intervalPoints(r.startSec, r.endSec);
          if (seg.length > 1 && seg[0].distance != null && seg[seg.length - 1].distance != null) {
            return sum + Math.max(0, seg[seg.length - 1].distance - seg[0].distance);
          }
          return sum;
        }, 0);
        const selectedPts = pts.filter((p) => !inCut(p.t) && validRanges.some((r) => p.t >= r.startSec && p.t <= r.endSec));
        const activeCuts = normalizeCuts([...(analyzeState.workCuts || []), ...(analyzeState.pendingCuts || [])]);
        const effectiveTotal = Math.max(1, totalSec - activeCuts.reduce((sum, c) => sum + Math.max(0, c.endSec - c.startSec), 0));
        const frac = Math.max(0, duration) / effectiveTotal;
        const totalTss = Number(summary.tss || data.tss_override || activityToTss(data) || 0);
        const tss = totalTss > 0 ? (totalTss * frac) : null;
        const minVal = (k) => {
          const vals = selectedPts.map((p) => p[k]).filter((v) => v !== null);
          return vals.length ? Math.min(...vals) : null;
        };
        const mean = (k) => {
          const vals = selectedPts.map((p) => p[k]).filter((v) => v !== null);
          return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
        };
        const maxVal = (k) => {
          const vals = selectedPts.map((p) => p[k]).filter((v) => v !== null);
          return vals.length ? Math.max(...vals) : null;
        };
        return { duration: Math.max(0, duration), distance, tss, minVal, mean, maxVal };
      }

      function renderSelectionStats() {
        const sel = analyzeState.selection;
        const hasLapSelection = analyzeState.selectionMode === 'laps' && Array.isArray(analyzeState.lapHighlightRanges) && analyzeState.lapHighlightRanges.length > 0;
        if (selectionTitle) selectionTitle.textContent = (sel || hasLapSelection) ? 'Selection' : 'Entire Workout';
        let stats;
        if (sel && analyzeState.selectionMode === 'graph') {
          const startT = Math.min(sel.start, sel.end);
          const endT = Math.max(sel.start, sel.end);
          stats = statsForRange(startT, endT);
        } else if (hasLapSelection) {
          stats = statsForDiscreteRanges(analyzeState.lapHighlightRanges);
        } else {
          const activeCuts = normalizeCuts([...(analyzeState.workCuts || []), ...(analyzeState.pendingCuts || [])]);
          const kept = [];
          let cursor = 0;
          activeCuts.forEach((c) => {
            if (c.startSec > cursor) kept.push({ startSec: cursor, endSec: c.startSec });
            cursor = Math.max(cursor, c.endSec);
          });
          if (cursor < totalSec) kept.push({ startSec: cursor, endSec: totalSec });
          stats = statsForDiscreteRanges(kept);
        }
        const {
          duration, distance, tss, minVal, mean, maxVal,
        } = stats;
        const metricRows = [];
        const pushRow = (name, key, unit, convert = (v) => v) => {
          const min = minVal(key);
          const avg = mean(key);
          const max = maxVal(key);
          if (avg == null && min == null && max == null) return;
          metricRows.push(`
            <tr>
              <td>${name}</td>
              <td>${min == null ? '' : Math.round(convert(min))}</td>
              <td>${avg == null ? '' : Math.round(convert(avg))}</td>
              <td>${max == null ? '' : Math.round(convert(max))}</td>
              <td>${unit}</td>
            </tr>
          `);
        };
        pushRow('Power', 'power', 'W');
        pushRow('Heart Rate', 'heart_rate', 'BPM');
        pushRow('Cadence', 'cadence', 'RPM');
        pushRow('Speed', 'speed', speedUnitLabel, toDisplaySpeed);
        const paceAvg = mean('speed');
        if (paceAvg != null && paceAvg > 0) {
          metricRows.push(`
            <tr>
              <td>Pace</td>
              <td></td>
              <td>${toDisplayPace(paceAvg) || ''}</td>
              <td></td>
              <td>${paceUnitLabel}</td>
            </tr>
          `);
        }

        statsNode.innerHTML = `
          <div class="selection-top-row">
            <div class="selection-top-cell">Duration<strong>${hms(duration)}</strong></div>
            <div class="selection-top-cell">Distance<strong>${fmtDistanceMeters(distance)}</strong></div>
            <div class="selection-top-cell">TSS<strong>${tss == null ? '' : Math.round(tss)}</strong></div>
          </div>
          <table class="selection-metrics">
            <thead>
              <tr><th>Metric</th><th>Min</th><th>Avg</th><th>Max</th><th>Unit</th></tr>
            </thead>
            <tbody>${metricRows.join('')}</tbody>
          </table>
        `;
      }

      function renderChannelList() {
        channelNode.innerHTML = '';
        const closeChannelPopup = () => {
          if (activeChannelPopup) {
            activeChannelPopup.remove();
            activeChannelPopup = null;
          }
        };
        lineMeta
          .filter((m) => existsChannel(m.key) && !analyzeState.workDeleted.has(m.key))
          .forEach((m) => {
          const isHidden = analyzeState.hiddenChannels.has(m.key);
          const item = document.createElement('button');
          item.type = 'button';
          const activeCls = !isHidden ? ' active' : '';
          item.className = `analyze-channel-item${activeCls}`;
          item.innerHTML = `
            <span class="channel-name${isHidden ? ' hidden-ch' : ''}">${m.label}</span>
          `;
          item.style.background = isHidden ? '' : m.color;
          item.style.borderColor = isHidden ? '' : m.color;
          item.addEventListener('click', (ev) => {
            ev.stopPropagation();
            closeChannelPopup();
            const popup = document.createElement('div');
            popup.className = 'channel-popup';
            const allVisible = lineMeta
              .filter((x) => existsChannel(x.key) && !analyzeState.workDeleted.has(x.key))
              .every((x) => !analyzeState.hiddenChannels.has(x.key));
            const actions = [];
            if (!allVisible) actions.push('<button data-action="show-all">Show All</button>');
            actions.push(`<button data-action="${isHidden ? 'show' : 'hide'}">${isHidden ? 'Show' : 'Hide'}</button>`);
            actions.push('<button data-action="hide-others">Hide Others</button>');
            actions.push('<button data-action="delete" class="ch-danger">Delete</button>');
            popup.innerHTML = `
              ${actions.join('')}
            `;
            popup.querySelectorAll('button').forEach((btn) => {
              btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const action = btn.dataset.action;
                if (action === 'show-all') {
                  analyzeState.hiddenChannels.clear();
                  analyzeState.workDeleted.clear();
                } else if (action === 'show') {
                  analyzeState.hiddenChannels.delete(m.key);
                } else if (action === 'hide') {
                  const currentlyVisible = visibleChannels().length;
                  if (currentlyVisible > 1) {
                    analyzeState.hiddenChannels.add(m.key);
                  } else {
                    analyzeState.hiddenChannels.clear();
                  }
                } else if (action === 'hide-others') {
                  analyzeState.hiddenChannels.clear();
                  lineMeta.forEach((c) => {
                    if (c.key !== m.key && !analyzeState.workDeleted.has(c.key) && existsChannel(c.key)) {
                      analyzeState.hiddenChannels.add(c.key);
                    }
                  });
                } else if (action === 'delete') {
                  analyzeState.workDeleted.add(m.key);
                  analyzeState.hiddenChannels.delete(m.key);
                }
                syncPendingState();
                popup.remove();
                renderChannelList();
                renderMain();
              });
            });
            document.body.appendChild(popup);
            const rect = item.getBoundingClientRect();
            const x = rect.left + (rect.width / 2) - (popup.offsetWidth / 2);
            const y = rect.top - popup.offsetHeight - 8;
            popup.style.left = `${Math.max(6, x)}px`;
            popup.style.top = `${Math.max(6, y)}px`;
            activeChannelPopup = popup;
            const close = (evt) => {
              if (!popup.contains(evt.target) && !item.contains(evt.target)) {
                closeChannelPopup();
                document.removeEventListener('click', close);
              }
            };
            setTimeout(() => document.addEventListener('click', close), 0);
          });
          channelNode.appendChild(item);
        });
      }

      function renderMain() {
        const xTicks = 6;
        const range = Math.max(1, analyzeState.wEnd - analyzeState.wStart);
        let svg = `<rect x="0" y="0" width="${w}" height="${h}" fill="#f3f7fd" stroke="#d6e1ee" rx="4"/>`;
        svg += `<rect x="${left}" y="${top}" width="${cw}" height="${ch}" fill="none" stroke="#d7dde6" stroke-width="0.9"/>`;
        for (let i = 0; i <= xTicks; i += 1) {
          const x = left + (i / xTicks) * cw;
          const sec = analyzeState.wStart + (i / xTicks) * range;
          svg += `<line x1="${x}" y1="${top}" x2="${x}" y2="${top + ch}" stroke="#dde8f4" stroke-width="1"/>`;
          svg += `<text x="${x}" y="${h - 8}" fill="#6b8099" font-size="10" text-anchor="middle">${hms(sec)}</text>`;
        }

        const tickVals = (min, max, count = 5) => {
          const span = Math.max(1e-6, max - min);
          const out = [];
          for (let i = 0; i <= count; i += 1) out.push(min + ((count - i) / count) * span);
          return out;
        };
        const axisForChannel = (meta) => {
          const smooth = smoothedValues(meta.key);
          const vals = [];
          for (let i = 0; i < pts.length; i += 1) {
            if (inCut(pts[i].t)) continue;
            const dT = displayFromOriginal(pts[i].t);
            if (dT < analyzeState.wStart || dT > analyzeState.wEnd) continue;
            const v = valueForAxis(meta.key, smooth[i]);
            if (v !== null) vals.push(v);
          }
          if (!vals.length) return null;
          const min = Math.min(...vals);
          const max = Math.max(...vals);
          const pad = Math.max(1, (max - min) * 0.06);
          return { meta, min: min - pad, max: max + pad, ticks: tickVals(min - pad, max + pad, 5) };
        };
        const axisOffsets = {
          cadence: { x: left - 8, tickX1: left, tickX2: left - 5, anchor: 'end' },
          power: { x: left - 34, tickX1: left, tickX2: left - 5, anchor: 'end' },
          heart_rate: { x: left + cw + 8, tickX1: left + cw, tickX2: left + cw + 5, anchor: 'start' },
          speed: { x: left + cw + 34, tickX1: left + cw, tickX2: left + cw + 5, anchor: 'start' },
        };
        visibleChannels().forEach((meta) => {
          const axis = axisForChannel(meta);
          if (!axis || !axisOffsets[meta.key]) return;
          const off = axisOffsets[meta.key];
          axis.ticks.forEach((t) => {
            const y = top + ((axis.max - t) / Math.max(1e-6, axis.max - axis.min)) * ch;
            svg += `<line x1="${off.tickX1}" y1="${y}" x2="${off.tickX2}" y2="${y}" stroke="#d7dde6" stroke-width="0.7"/>`;
            svg += `<text x="${off.x}" y="${y + 2.5}" fill="${meta.color}" font-size="7.5" text-anchor="${off.anchor}">${Math.round(t)}</text>`;
          });
        });

        if (analyzeState.selectionMode === 'laps' && Array.isArray(analyzeState.lapHighlightRanges)) {
          analyzeState.lapHighlightRanges.forEach((rangeSel) => {
            const dStart = displayFromOriginal(rangeSel.startSec);
            const dEnd = displayFromOriginal(rangeSel.endSec);
            const rStart = Math.max(analyzeState.wStart, Math.min(dStart, dEnd));
            const rEnd = Math.min(analyzeState.wEnd, Math.max(dStart, dEnd));
            if (rEnd <= rStart) return;
            const sx1 = timeToX(rStart);
            const sx2 = timeToX(rEnd);
            svg += `<rect x="${sx1.toFixed(1)}" y="${top}" width="${Math.max(1, sx2 - sx1).toFixed(1)}" height="${ch}" fill="rgba(50,120,255,0.12)" stroke="rgba(40,100,240,0.35)" stroke-width="1"/>`;
          });
        } else if (analyzeState.selectionMode === 'graph' && analyzeState.selection) {
          const sx1 = timeToX(Math.min(analyzeState.selection.start, analyzeState.selection.end));
          const sx2 = timeToX(Math.max(analyzeState.selection.start, analyzeState.selection.end));
          svg += `<rect x="${sx1.toFixed(1)}" y="${top}" width="${Math.max(1, sx2 - sx1).toFixed(1)}" height="${ch}" fill="rgba(50,120,255,0.12)" stroke="rgba(40,100,240,0.35)" stroke-width="1"/>`;
        }

        visibleChannels().forEach((meta) => {
          const smooth = smoothedValues(meta.key);
          const win = [];
          for (let i = 0; i < pts.length; i += 1) {
            if (inCut(pts[i].t)) continue;
            const dT = displayFromOriginal(pts[i].t);
            if (dT >= analyzeState.wStart && dT <= analyzeState.wEnd && smooth[i] !== null) {
              win.push({ t: dT, origT: pts[i].t, v: smooth[i] });
            }
          }
          if (!win.length) return;
          const min = Math.min(...win.map((x) => x.v));
          const max = Math.max(...win.map((x) => x.v));
          const span = Math.max(0.001, max - min);
          let d = '';
          let prevOrig = null;
          win.forEach((p, idx) => {
            const x = timeToX(p.t);
            const y = top + (1 - ((p.v - min) / span)) * ch;
            const allCuts = normalizeCuts([...(analyzeState.workCuts || []), ...(analyzeState.pendingCuts || [])]);
            const cutBetween = prevOrig !== null && allCuts.some((c) => c.startSec <= p.origT && c.endSec >= prevOrig);
            const breakSeg = prevOrig !== null && ((p.origT - prevOrig) > 8 || cutBetween);
            d += `${idx && !breakSeg ? 'L' : 'M'}${x.toFixed(2)} ${y.toFixed(2)} `;
            prevOrig = p.origT;
          });
          svg += `<path d="${d.trim()}" stroke="${meta.color}" stroke-width="0.75" opacity="0.95" fill="none" stroke-linecap="butt" stroke-linejoin="miter" shape-rendering="geometricPrecision"/>`;
        });

        chart.innerHTML = svg;
        if (analyzeState.cursorSvgX >= 0) {
          const ns = 'http://www.w3.org/2000/svg';
          const line = document.createElementNS(ns, 'line');
          line.setAttribute('x1', String(analyzeState.cursorSvgX));
          line.setAttribute('x2', String(analyzeState.cursorSvgX));
          line.setAttribute('y1', String(top));
          line.setAttribute('y2', String(top + ch));
          line.setAttribute('stroke', '#a7b7ca');
          line.setAttribute('stroke-width', '1');
          line.setAttribute('opacity', '0.9');
          chart.appendChild(line);
        }
        const eff = effectivePoints();
        const effectiveTotal = Math.max(1, (eff.length ? eff[eff.length - 1].dT : totalSec));
        const canZoom = !!(analyzeState.selectionMode === 'graph' && analyzeState.selection && Math.abs(analyzeState.selection.end - analyzeState.selection.start) > 1);
        const hasLapCutSelection = analyzeState.selectionMode === 'laps'
          && Array.isArray(analyzeState.lapHighlightRanges)
          && analyzeState.lapHighlightRanges.some((r) => Number(r.endSec || 0) - Number(r.startSec || 0) > 1);
        const isZoomed = analyzeState.wStart > 0 || analyzeState.wEnd < effectiveTotal;
        zoomBtn.textContent = isZoomed && !canZoom ? 'Unzoom' : 'Zoom';
        zoomBtn.disabled = !canZoom && !isZoomed;
        cutBtn.disabled = !(canZoom || hasLapCutSelection);
        if (unzoomBtn) unzoomBtn.classList.add('hidden');
        renderSelectionStats();
      }

      function hideCursorTooltip() {
        analyzeState.cursorSvgX = -1;
        const tooltip = document.getElementById('wvCursorTooltip');
        tooltip.classList.add('hidden');
      }

      function updateCursor(ev) {
        const { time, svgX } = plotMouseToTime(ev);
        analyzeState.cursorSvgX = svgX;
        const nearest = findNearestPointAt(time);
        const tooltip = document.getElementById('wvCursorTooltip');
        let html = `<strong>Time:</strong> ${hms(time)}`;
        if (nearest.cadence != null) html += `<br/><strong>RPM:</strong> ${Math.round(nearest.cadence)} rpm`;
        if (nearest.heart_rate != null) html += `<br/><strong>BPM:</strong> ${Math.round(nearest.heart_rate)} bpm`;
        if (nearest.power != null) html += `<br/><strong>Watts:</strong> ${Math.round(nearest.power)} W`;
        if (nearest.speed != null) html += `<br/><strong>Speed:</strong> ${fmtAxis(nearest.speed, 'speed')}`;
        tooltip.innerHTML = html;
        tooltip.classList.remove('hidden');
        const box = chart.closest('.chart-box').getBoundingClientRect();
        let tx = ev.clientX - box.left + 12;
        let ty = ev.clientY - box.top - 12;
        if (tx + 170 > box.width) tx -= 180;
        if (ty < 2) ty = 2;
        tooltip.style.left = `${tx}px`;
        tooltip.style.top = `${ty}px`;
      }

      let dragStart = null;
      chart.onmousedown = (ev) => {
        if (ev.button !== 0) return;
        ev.preventDefault();
        document.body.classList.add('analyze-dragging');
        analyzeState.selectedLapKeys.clear();
        analyzeState.lapHighlightRanges = [];
        lapBody.querySelectorAll('.lap-select').forEach((cb) => { cb.checked = false; });
        lapBody.querySelectorAll('tr').forEach((row) => row.classList.remove('selected'));
        const { time } = plotMouseToTime(ev);
        dragStart = time;
        analyzeState.selectionMode = 'graph';
        analyzeState.selection = { start: time, end: time };
      };
      chart.onmousemove = (ev) => {
        updateCursor(ev);
        if (dragStart !== null && (ev.buttons & 1) === 0) {
          dragStart = null;
          document.body.classList.remove('analyze-dragging');
          return;
        }
        if (dragStart === null) return;
        const { time } = plotMouseToTime(ev);
        analyzeState.selectionMode = 'graph';
        analyzeState.selection = { start: dragStart, end: time };
        renderMain();
      };
      chart.onmouseup = (ev) => {
        if (dragStart === null) return;
        const { time } = plotMouseToTime(ev);
        analyzeState.selectionMode = 'graph';
        analyzeState.selection = { start: dragStart, end: time };
        if (Math.abs(time - dragStart) <= 1) {
          analyzeState.selection = null;
          analyzeState.selectionMode = 'none';
        }
        dragStart = null;
        document.body.classList.remove('analyze-dragging');
        renderMain();
      };
      chart.onmouseleave = () => {
        if (dragStart !== null) return;
        hideCursorTooltip();
        renderMain();
      };

      zoomBtn.onclick = () => {
        const eff = effectivePoints();
        const effectiveTotal = Math.max(1, (eff.length ? eff[eff.length - 1].dT : totalSec));
        const isZoomed = analyzeState.wStart > 0 || analyzeState.wEnd < effectiveTotal;
        const canZoom = !!(analyzeState.selectionMode === 'graph' && analyzeState.selection && Math.abs(analyzeState.selection.end - analyzeState.selection.start) > 1);
        if (isZoomed && !canZoom) {
          analyzeState.zoomStack = [];
          analyzeState.wStart = 0;
          analyzeState.wEnd = effectiveTotal;
          analyzeState.selection = null;
          analyzeState.selectionMode = 'none';
          renderMain();
          return;
        }
        if (analyzeState.selectionMode !== 'graph' || !analyzeState.selection) return;
        const start = Math.min(analyzeState.selection.start, analyzeState.selection.end);
        const end = Math.max(analyzeState.selection.start, analyzeState.selection.end);
        if (Math.abs(end - start) <= 1) return;
        analyzeState.zoomStack.push({ s: analyzeState.wStart, e: analyzeState.wEnd });
        analyzeState.wStart = start;
        analyzeState.wEnd = end;
        analyzeState.selection = null;
        analyzeState.selectionMode = 'none';
        renderMain();
      };
      cutBtn.onclick = () => {
        const cutRanges = [];
        if (analyzeState.selectionMode === 'graph' && analyzeState.selection) {
          const dStart = Math.min(analyzeState.selection.start, analyzeState.selection.end);
          const dEnd = Math.max(analyzeState.selection.start, analyzeState.selection.end);
          if (Math.abs(dEnd - dStart) > 1) {
            cutRanges.push({ startSec: originalFromDisplay(dStart), endSec: originalFromDisplay(dEnd) });
          }
        } else if (analyzeState.selectionMode === 'laps' && Array.isArray(analyzeState.lapHighlightRanges)) {
          analyzeState.lapHighlightRanges.forEach((r) => {
            const start = Number(r.startSec || 0);
            const end = Number(r.endSec || 0);
            if ((end - start) > 1) cutRanges.push({ startSec: start, endSec: end });
          });
        }
        if (!cutRanges.length) return;
        analyzeState.pendingCuts.push(...cutRanges);
        analyzeState.selection = null;
        analyzeState.selectionMode = 'none';
        analyzeState.selectedLapKeys.clear();
        analyzeState.lapHighlightRanges = [];
        lapBody.querySelectorAll('.lap-select').forEach((cb) => { cb.checked = false; });
        lapBody.querySelectorAll('tr').forEach((row) => row.classList.remove('selected'));
        syncPendingState();
        renderMain();
      };
      if (unzoomBtn) {
        unzoomBtn.onclick = () => {
          analyzeState.zoomStack = [];
          analyzeState.wStart = 0;
          const eff = effectivePoints();
          analyzeState.wEnd = Math.max(1, (eff.length ? eff[eff.length - 1].dT : totalSec));
          analyzeState.selection = null;
          analyzeState.selectionMode = 'none';
          renderMain();
        };
      }
      smoothingStep.oninput = () => {
        analyzeState.smoothingStep = Number(smoothingStep.value || 2);
        analyzeState.smoothedCache.clear();
        renderMain();
      };
      smoothingStep.value = '2';

      if (applyBtn) {
        applyBtn.onclick = async () => {
          const ok = await confirmApplyChanges();
          if (!ok) return;
          await persistAnalyzeEdits();
        };
      }
      if (cancelBtn) {
        cancelBtn.onclick = () => {
          analyzeState.pendingCuts = [];
          analyzeState.selection = null;
          analyzeState.selectionMode = 'none';
          analyzeState.workDeleted = new Set(analyzeState.persistedDeleted);
          analyzeState.hiddenChannels.clear();
          syncPendingState();
          renderChannelList();
          renderMain();
        };
      }
      analyzeState.applyPending = async () => { await persistAnalyzeEdits(); };
      analyzeState.cancelPending = () => {
        analyzeState.pendingCuts = [];
        analyzeState.selection = null;
        analyzeState.selectionMode = 'none';
        analyzeState.workDeleted = new Set(analyzeState.persistedDeleted);
        analyzeState.hiddenChannels.clear();
        syncPendingState();
        renderChannelList();
        renderMain();
      };

      const toDisplay = (value, fn) => (value == null ? '' : fn(value));
      const lapRows = laps.length ? laps : [{
        name: 'Lap 1',
        start: series[0].timestamp,
        end: series[series.length - 1].timestamp,
        duration_s: totalSec,
      }];
      const ftpKey = activitySportKey({ type: (payload.planned && payload.planned.workout_type) || data.type || data.workout_type || 'ride' });
      const ftpValue = Number((appSettings.ftp || {})[ftpKey] || 0);
      function overlap(a1, a2, b1, b2) {
        return Math.max(0, Math.min(a2, b2) - Math.max(a1, b1));
      }
      function lapKeptDuration(startSec, endSec) {
        let kept = Math.max(0, endSec - startSec);
        const activeCuts = normalizeCuts([...(analyzeState.workCuts || []), ...(analyzeState.pendingCuts || [])]);
        activeCuts.forEach((c) => {
          kept -= overlap(startSec, endSec, c.startSec, c.endSec);
        });
        return Math.max(0, kept);
      }
      function applyLapSelectionFromChecks() {
        const checks = Array.from(lapBody.querySelectorAll('.lap-select'));
        const selected = checks
          .filter((cb) => cb.checked)
          .map((cb) => ({
            key: cb.dataset.lapKey,
            startSec: Number(cb.dataset.startSec || 0),
            endSec: Number(cb.dataset.endSec || 0),
            row: cb.closest('tr'),
          }))
          .filter((x) => x.endSec > x.startSec);
        analyzeState.selectedLapKeys = new Set(selected.map((x) => x.key));
        analyzeState.lapHighlightRanges = selected.map((x) => ({ startSec: x.startSec, endSec: x.endSec }));
        Array.from(lapBody.querySelectorAll('tr')).forEach((row) => row.classList.remove('selected'));
        selected.forEach((x) => {
          if (x.row) x.row.classList.add('selected');
        });
        if (!selected.length) {
          analyzeState.selection = null;
          analyzeState.selectionMode = 'none';
        } else {
          analyzeState.selection = null;
          analyzeState.selectionMode = 'laps';
        }
        renderMain();
      }
      function renderLapRows() {
        lapBody.innerHTML = '';
        lapRows.forEach((lap, idx) => {
          const startSec = Math.max(0, timeToSec(lap.start || series[0].timestamp, baseMs));
          const endSec = Math.max(startSec + 1, timeToSec(lap.end || series[series.length - 1].timestamp, baseMs));
          const keptPts = pts.filter((p) => p.t >= startSec && p.t <= endSec && !inCut(p.t, false));
          const dur = lapKeptDuration(startSec, endSec);
          if (dur <= 0) return;
          const ifVal = Number(lap.if || lap.intensity_factor || 0) || null;
          const dist = keptPts.length > 1 && keptPts[0].distance != null && keptPts[keptPts.length - 1].distance != null
            ? Math.max(0, keptPts[keptPts.length - 1].distance - keptPts[0].distance)
            : Number(lap.distance_m || 0);
          const avg = (k) => {
            const vals = keptPts.map((p) => p[k]).filter((v) => v != null);
            return vals.length ? (vals.reduce((s, v) => s + v, 0) / vals.length) : null;
          };
          const mx = (k) => {
            const vals = keptPts.map((p) => p[k]).filter((v) => v != null);
            return vals.length ? Math.max(...vals) : null;
          };
          const powerPts = keptPts.filter((p) => p.power != null).map((p) => ({ t: p.t, p: p.power }));
          let npLap = null;
          if (powerPts.length) {
            const window = [];
            let sum = 0;
            const rolling = [];
            for (let i = 0; i < powerPts.length; i += 1) {
              const pt = powerPts[i];
              window.push(pt);
              sum += pt.p;
              while (window.length && (pt.t - window[0].t) > 30) {
                sum -= window[0].p;
                window.shift();
              }
              if (window.length) rolling.push(sum / window.length);
            }
            if (rolling.length) {
              const mean4 = rolling.reduce((s, v) => s + (v ** 4), 0) / rolling.length;
              npLap = mean4 ** 0.25;
            }
          }
          const ifCalc = (npLap && ftpValue > 0) ? (npLap / ftpValue) : null;
          const tss = (npLap && ifCalc) ? ((dur * npLap * ifCalc) / (ftpValue * 3600) * 100) : null;
          const workCalc = (avg('power') != null && dur > 0) ? ((avg('power') * dur) / 1000) : null;
          const row = document.createElement('tr');
          const lapKey = String(idx);
          const checked = analyzeState.selectedLapKeys.has(lapKey) ? 'checked' : '';
          row.innerHTML = `
            <td class="lap-check-col"><input type="checkbox" class="lap-select" data-lap-key="${lapKey}" data-start-sec="${startSec}" data-end-sec="${endSec}" ${checked} /></td>
            <td>${lap.name || `Lap ${idx + 1}`}</td>
            <td>${fmtTimeShort(lap.start)}</td>
            <td>${fmtTimeShort(lap.end)}</td>
            <td>${toDisplay(dur, hms)}</td>
            <td>${toDisplay(dur, hms)}</td>
            <td>${toDisplay(dist, fmtDistanceMeters)}</td>
            <td>${tss == null ? '' : Math.round(tss)}</td>
            <td>${ifCalc == null ? '' : ifCalc.toFixed(2)}</td>
            <td>${npLap == null ? '' : Math.round(npLap)}</td>
            <td>${avg('power') == null ? '' : Math.round(avg('power'))}</td>
            <td>${mx('power') == null ? '' : Math.round(mx('power'))}</td>
            <td>${avg('heart_rate') == null ? '' : Math.round(avg('heart_rate'))}</td>
            <td>${mx('heart_rate') == null ? '' : Math.round(mx('heart_rate'))}</td>
            <td>${avg('speed') == null ? '' : fmtAxis(avg('speed'), 'speed')}</td>
            <td>${mx('speed') == null ? '' : fmtAxis(mx('speed'), 'speed')}</td>
            <td>${avg('cadence') == null ? '' : Math.round(avg('cadence'))}</td>
            <td>${workCalc == null ? '' : Math.round(workCalc)}</td>
            <td>${lap.calories == null ? '' : lap.calories}</td>
          `;
          const check = row.querySelector('.lap-select');
          if (check) {
            check.addEventListener('change', () => {
              applyLapSelectionFromChecks();
            });
          }
          lapBody.appendChild(row);
        });
        applyLapSelectionFromChecks();
      }
      renderLapRows();

      document.getElementById('wvHrMin').value = summary.min_hr ? String(Math.round(summary.min_hr)) : '';
      document.getElementById('wvHrAvg').value = summary.avg_hr ? String(Math.round(summary.avg_hr)) : '';
      document.getElementById('wvHrMax').value = summary.max_hr ? String(Math.round(summary.max_hr)) : '';
      document.getElementById('wvPowerMin').value = summary.min_power ? String(Math.round(summary.min_power)) : '';
      document.getElementById('wvPowerAvg').value = summary.avg_power ? String(Math.round(summary.avg_power)) : '';
      document.getElementById('wvPowerMax').value = summary.max_power ? String(Math.round(summary.max_power)) : '';
      renderChannelList();
      syncPendingState();
      renderMain();
    }

    function buildCtlProjection(eventDateKey) {
      const todayMet = buildMetricsToDate(todayKey());
      const historySeries = todayMet.ctlSeries.slice(-90);
      const todayCtl = historySeries[historySeries.length - 1];

      const plannedMap = {};
      const today = todayKey();
      calendarItems
        .filter(i => i.kind === 'workout' && i.date > today)
        .forEach(i => { plannedMap[i.date] = (plannedMap[i.date] || 0) + itemToTss(i); });

      const projectedSeries = [todayCtl];
      let ctlPrev = todayCtl;
      const cur = parseDateKey(today);
      const end = parseDateKey(eventDateKey);
      cur.setDate(cur.getDate() + 1);
      while (cur <= end) {
        const tss = plannedMap[dateKeyFromDate(cur)] || 0;
        ctlPrev = ctlPrev + (tss - ctlPrev) / 42;
        projectedSeries.push(ctlPrev);
        cur.setDate(cur.getDate() + 1);
      }

      return { historySeries, projectedSeries, todayCtl, eventCtl: ctlPrev };
    }

    function renderEventCtlChart(container, event) {
      const proj = buildCtlProjection(event.date);
      const { historySeries, projectedSeries, todayCtl, eventCtl } = proj;

      const W = 500, H = 160, PAD = { t: 16, r: 80, b: 20, l: 10 };
      const chartW = W - PAD.l - PAD.r;
      const chartH = H - PAD.t - PAD.b;
      const totalPoints = historySeries.length + projectedSeries.length - 1;
      const allVals = [...historySeries, ...projectedSeries.slice(1)];
      const minV = Math.min(...allVals) * 0.92;
      const maxV = Math.max(...allVals, 1) * 1.06;

      const xOf = i => PAD.l + (i / (totalPoints - 1)) * chartW;
      const yOf = v => PAD.t + chartH - ((v - minV) / (maxV - minV)) * chartH;

      const histPts = historySeries.map((v, i) => `${xOf(i).toFixed(1)},${yOf(v).toFixed(1)}`).join(' ');
      const todayIdx = historySeries.length - 1;
      const projPts = projectedSeries.map((v, i) => `${xOf(todayIdx + i).toFixed(1)},${yOf(v).toFixed(1)}`).join(' ');

      const todayX = xOf(todayIdx), todayY = yOf(todayCtl);
      const eventX = xOf(totalPoints - 1), eventY = yOf(eventCtl);
      const eventCtlRounded = Math.round(eventCtl);
      const todayCtlRounded = Math.round(todayCtl);

      container.innerHTML = `
        <svg class="event-ctl-chart" viewBox="0 0 ${W} ${H}" width="100%" height="${H}" aria-label="CTL trend to event">
          <defs>
            <linearGradient id="evHistGrad" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stop-color="#5a8fd4" stop-opacity="0.18"/>
              <stop offset="100%" stop-color="#5a8fd4" stop-opacity="0"/>
            </linearGradient>
          </defs>
          <polyline points="${histPts}" fill="none" stroke="#4a7fc1" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
          <polyline points="${projPts}" fill="none" stroke="#6a9fd8" stroke-width="1.8" stroke-dasharray="5,4" stroke-linejoin="round" stroke-linecap="round"/>
          <circle cx="${todayX.toFixed(1)}" cy="${todayY.toFixed(1)}" r="5" fill="#4a7fc1" stroke="#fff" stroke-width="1.5"/>
          <text x="${(todayX + 8).toFixed(1)}" y="${(todayY - 6).toFixed(1)}" class="ectl-label" font-weight="700">Today ${todayCtlRounded} CTL</text>
          <circle cx="${eventX.toFixed(1)}" cy="${eventY.toFixed(1)}" r="5" fill="#7a9fc0" stroke="#fff" stroke-width="1.5"/>
          <text x="${(eventX - 4).toFixed(1)}" y="${(eventY - 10).toFixed(1)}" class="ectl-label" text-anchor="end">Event ${eventCtlRounded} CTL</text>
          <text x="${(eventX - 4).toFixed(1)}" y="${(eventY - 22).toFixed(1)}" class="ectl-label" text-anchor="end">${event.title}</text>
        </svg>`;
    }

    function renderEvents() {
      const list = document.getElementById('eventsList');
      const today = todayKey();
      const allEvents = calendarItems
        .filter(i => i.kind === 'event')
        .sort((a, b) => (a.date > b.date ? 1 : -1));

      list.innerHTML = '';
      if (!allEvents.length) {
        list.innerHTML = '<p class="meta">No events yet. Click + to add one.</p>';
        return;
      }

      const priorityOrder = { A: 0, B: 1, C: 2 };
      const upcoming = allEvents.filter(e => e.date >= today);
      upcoming.sort((a, b) => {
        const pa = priorityOrder[a.priority] ?? 2;
        const pb = priorityOrder[b.priority] ?? 2;
        if (pa !== pb) return pa - pb;
        return a.date > b.date ? 1 : -1;
      });
      const featured = upcoming[0] || allEvents[allEvents.length - 1];
      const eventDate = parseDateKey(featured.date);
      const todayDate = parseDateKey(today);
      const daysUntil = Math.round((eventDate - todayDate) / 86400000);
      const weeksUntil = Math.round(daysUntil / 7);
      const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
      const mon = months[eventDate.getMonth()];
      const day = eventDate.getDate();

      const featuredEl = document.createElement('div');
      featuredEl.className = 'event-featured';

      const badgeEl = document.createElement('div');
      badgeEl.className = 'event-date-badge';
      badgeEl.innerHTML = `<span class="edb-month">${mon}</span><span class="edb-day">${day}</span>`;

      const infoEl = document.createElement('div');
      infoEl.className = 'event-featured-info';
      const countdownText = daysUntil > 0
        ? `${weeksUntil > 0 ? weeksUntil + ' WEEK' + (weeksUntil !== 1 ? 'S' : '') : daysUntil + ' DAY' + (daysUntil !== 1 ? 'S' : '')} UNTIL EVENT`
        : daysUntil === 0 ? 'TODAY' : 'PAST EVENT';
      infoEl.innerHTML = `<div class="event-countdown">${countdownText}</div>
        <div class="event-featured-title">${featured.title}</div>`;

      featuredEl.appendChild(badgeEl);
      featuredEl.appendChild(infoEl);
      list.appendChild(featuredEl);

      const chartDiv = document.createElement('div');
      chartDiv.className = 'event-ctl-chart-wrap';
      renderEventCtlChart(chartDiv, featured);
      list.appendChild(chartDiv);

      const tableDiv = document.createElement('div');
      tableDiv.className = 'events-table';
      allEvents.forEach((e, idx) => {
        const eDate = parseDateKey(e.date);
        const mo = months[eDate.getMonth()];
        const dy = String(eDate.getDate()).padStart(2, '0');
        const row = document.createElement('div');
        row.className = 'events-row' + (idx % 2 === 1 ? ' events-row-alt' : '');
        row.innerHTML = `<span class="events-row-date">${mo} ${dy}</span><span class="events-row-title">${e.title}</span>`;
        row.style.cursor = 'pointer';
        row.addEventListener('click', () => openDetailModal(e));
        tableDiv.appendChild(row);
      });
      list.appendChild(tableDiv);
    }

    function renderGoals() {
      const list = document.getElementById('goalsList');
      const goals = calendarItems
        .filter(i => i.kind === 'goal')
        .sort((a, b) => {
          const so = (a.sort_order ?? 0) - (b.sort_order ?? 0);
          if (so !== 0) return so;
          return a.date > b.date ? 1 : -1;
        });

      list.innerHTML = '';
      if (!goals.length) {
        list.innerHTML = '<p class="meta">No goals yet. Click Add Goal.</p>';
        return;
      }

      const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

      // Header row
      const header = document.createElement('div');
      header.className = 'goals-header-row';
      header.innerHTML = '<span class="goals-col-date">Date</span><span class="goals-col-label">Goals</span>';
      list.appendChild(header);

      let dragSrc = null;

      goals.forEach((g, idx) => {
        const gDate = parseDateKey(g.date);
        const dateLabel = `${months[gDate.getMonth()]} ${gDate.getDate()}`;

        const row = document.createElement('div');
        row.className = 'goals-row' + (idx % 2 === 1 ? ' goals-row-alt' : '') + (g.completed ? ' goals-row-done' : '');
        row.draggable = true;
        row.dataset.id = g.id;

        row.innerHTML = `
          <span class="goals-col-date goals-date-cell">${dateLabel}</span>
          <span class="goals-col-goal">
            <span class="goals-drag-handle" title="Drag to reorder">⠿</span>
            <input type="checkbox" class="goals-check" ${g.completed ? 'checked' : ''} aria-label="Mark complete" />
            <span class="goals-title${g.completed ? ' goals-title-done' : ''}">${g.title}</span>
          </span>`;

        // Checkbox toggle
        row.querySelector('.goals-check').addEventListener('change', async (e) => {
          e.stopPropagation();
          await fetch(`/calendar-items/${g.id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ...g, completed: e.target.checked }),
          });
          await loadData();
        });

        // Click row to open edit (not on checkbox/handle)
        row.addEventListener('click', (e) => {
          if (e.target.closest('.goals-check') || e.target.closest('.goals-drag-handle')) return;
          openDetailModal(g);
        });

        // Drag-and-drop
        row.addEventListener('dragstart', (e) => {
          dragSrc = row;
          e.dataTransfer.effectAllowed = 'move';
          row.classList.add('goals-dragging');
        });
        row.addEventListener('dragend', () => {
          row.classList.remove('goals-dragging');
          list.querySelectorAll('.goals-row').forEach(r => r.classList.remove('goals-drag-over'));
        });
        row.addEventListener('dragover', (e) => {
          e.preventDefault();
          e.dataTransfer.dropEffect = 'move';
          if (row !== dragSrc) {
            list.querySelectorAll('.goals-row').forEach(r => r.classList.remove('goals-drag-over'));
            row.classList.add('goals-drag-over');
          }
        });
        row.addEventListener('drop', async (e) => {
          e.preventDefault();
          if (!dragSrc || dragSrc === row) return;
          row.classList.remove('goals-drag-over');

          // Reorder in DOM to determine new sequence
          const rows = [...list.querySelectorAll('.goals-row')];
          const srcIdx = rows.indexOf(dragSrc);
          const tgtIdx = rows.indexOf(row);
          if (srcIdx < tgtIdx) row.after(dragSrc); else row.before(dragSrc);

          // Persist new sort_order for all goals
          const reordered = [...list.querySelectorAll('.goals-row')];
          await Promise.all(reordered.map((r, i) => {
            const item = calendarItems.find(ci => ci.id === r.dataset.id);
            if (!item) return Promise.resolve();
            return fetch(`/calendar-items/${item.id}`, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ ...item, sort_order: i }),
            });
          }));
          await loadData();
        });

        list.appendChild(row);
      });
    }

    function sportKeyToLabel(key) {
      return { ride: 'Cycling', run: 'Running', swim: 'Swimming', row: 'Rowing', strength: 'Strength', other: 'Other' }[key] || key;
    }

    function workoutTypeSportKey(type) {
      const t = String(type || '').toLowerCase();
      if (t.includes('ride') || t.includes('bike') || t.includes('cycl')) return 'ride';
      if (t.includes('run') || t.includes('walk')) return 'run';
      if (t.includes('swim')) return 'swim';
      if (t.includes('row')) return 'row';
      if (t.includes('strength') || t.includes('weight')) return 'strength';
      return 'other';
    }

    function buildPowerZones(ftp, sportLabel) {
      const zones = [
        { num: '6', lo: 1.21, hi: null, color: '#3d1a6e' },
        { num: '5', lo: 1.06, hi: 1.20, color: '#6b2090' },
        { num: '4', lo: 0.91, hi: 1.05, color: '#1e4fc2' },
        { num: '3', lo: 0.76, hi: 0.90, color: '#1a7fc4' },
        { num: '2', lo: 0.56, hi: 0.75, color: '#1a9ec4' },
        { num: '1', lo: 0,    hi: 0.55, color: '#b8d8eb' },
      ];
      const el = document.createElement('div');
      el.className = 'zones-block';
      el.innerHTML = `
        <div class="zones-title zones-title-power">Power: ${sportLabel}</div>
        <div class="zones-threshold">Threshold: ${ftp} W</div>
        <div class="zones-rows">
          ${zones.map(z => {
            const lo = Math.round(z.lo * ftp);
            const hi = z.hi ? Math.round(z.hi * ftp) : 2000;
            return `<div class="zone-row">
              <span class="zone-num">${z.num}</span>
              <span class="zone-bar-pip" style="background:${z.color}"></span>
              <span class="zone-range">${lo}-${hi}</span>
            </div>`;
          }).join('')}
        </div>`;
      return el;
    }

    function buildHrZones(lthr) {
      const zones = [
        { name: 'Zone 5C: Anaerobic Capacity', lo: 1.06, hi: null,  color: '#3d1a6e' },
        { name: 'Zone 5B: Aerobic Capacity',   lo: 1.03, hi: 1.06,  color: '#6b2090' },
        { name: 'Zone 5A: SuperThreshold',     lo: 1.00, hi: 1.02,  color: '#8b40a0' },
        { name: 'Zone 4: SubThreshold',        lo: 0.94, hi: 0.99,  color: '#1e4fc2' },
        { name: 'Zone 3: Tempo',               lo: 0.89, hi: 0.93,  color: '#1a7fc4' },
        { name: 'Zone 2: Aerobic',             lo: 0.81, hi: 0.88,  color: '#1a9ec4' },
        { name: 'Zone 1: Recovery',            lo: 0,    hi: 0.80,  color: '#b8d8eb' },
      ];
      const el = document.createElement('div');
      el.className = 'zones-block zones-block-hr';
      el.innerHTML = `
        <div class="zones-title zones-title-hr">Heart Rate</div>
        <div class="zones-threshold">Threshold: ${lthr} bpm</div>
        <div class="zones-rows">
          ${zones.map(z => {
            const lo = z.lo === 0 ? 0 : Math.round(z.lo * lthr);
            const hi = z.hi ? Math.round(z.hi * lthr) : 255;
            return `<div class="zone-row zone-row-hr">
              <span class="zone-name">${z.name}</span>
              <span class="zone-bar-pip" style="background:${z.color}"></span>
              <span class="zone-range">${lo}-${hi}</span>
            </div>`;
          }).join('')}
        </div>`;
      return el;
    }

    function renderTrainingZones(doneToday, plannedToday) {
      const panel = document.getElementById('trainingZonesPanel');
      const grid  = document.getElementById('trainingZonesGrid');
      const ftp   = appSettings.ftp || {};
      const lthr  = appSettings.lthr || {};

      // Collect relevant sport keys from today's activity
      const sportKeys = new Set();
      doneToday.forEach(a => sportKeys.add(activitySportKey(a)));
      plannedToday.forEach(p => sportKeys.add(workoutTypeSportKey(p.workout_type)));
      if (!sportKeys.size) Object.entries(ftp).forEach(([k, v]) => { if (v) sportKeys.add(k); });

      const anyFtp  = [...sportKeys].some(k => Number(ftp[k] || 0) > 0);
      const anyLthr = Number(lthr.global || lthr.run || lthr.ride || lthr.row || 0) > 0;
      if (!anyFtp && !anyLthr) { panel.style.display = 'none'; return; }
      panel.style.display = '';
      grid.innerHTML = '';

      const powerCol = document.createElement('div');
      powerCol.className = 'zones-col';
      const hrCol = document.createElement('div');
      hrCol.className = 'zones-col';

      sportKeys.forEach(key => {
        const ftpVal = Number(ftp[key] || 0);
        if (ftpVal > 0) powerCol.appendChild(buildPowerZones(ftpVal, sportKeyToLabel(key)));
      });

      const lthrVal = Number(lthr.global || lthr.run || lthr.ride || lthr.row || 0);
      if (lthrVal > 0) hrCol.appendChild(buildHrZones(lthrVal));

      const row = document.createElement('div');
      row.className = 'zones-grid';
      if (powerCol.children.length) row.appendChild(powerCol);
      if (hrCol.children.length) row.appendChild(hrCol);
      grid.appendChild(row);
    }

    function completionColorClass(actualMin, plannedMin) {
      if (!plannedMin || plannedMin <= 0) return '';
      const deviation = Math.abs(actualMin - plannedMin) / plannedMin;
      if (deviation <= 0.1667) return 'today-card-green';
      if (deviation <= 0.40)   return 'today-card-yellow';
      return 'today-card-orange';
    }

    function yesterdayKey() {
      const d = new Date();
      d.setDate(d.getDate() - 1);
      return dateKeyFromDate(d);
    }

    function tomorrowKey() {
      const d = new Date();
      d.setDate(d.getDate() + 1);
      return dateKeyFromDate(d);
    }

    function renderTomorrowFeed() {
      const tmrw = tomorrowKey();
      const feed = document.getElementById('tomorrowFeed');
      if (!feed) return;
      feed.innerHTML = '';
      const plannedTmrw = calendarItems.filter(i => i.kind === 'workout' && i.date === tmrw);
      if (!plannedTmrw.length) {
        feed.innerHTML = '<p class="meta" style="padding:10px 0;">No workouts planned.</p>';
        return;
      }
      plannedTmrw.forEach(p => {
        const sportKey = workoutTypeSportKey(p.workout_type);
        const iconSrc = ICON_ASSETS[sportKey] || ICON_ASSETS.other;
        const plannedMin = Number(p.duration_min || 0);
        const tss = itemToTss(p);
        const card = document.createElement('div');
        card.className = 'today-card today-card-planned';
        card.innerHTML = `
          <div class="today-card-sport today-card-sport-planned">
            ${p.workout_type || 'Workout'}
            <span class="badge planned" style="margin-left:auto;">Planned</span>
          </div>
          <div class="today-card-stats">
            <img class="today-card-icon" src="${iconSrc}" alt="${p.workout_type}" />
            <span class="today-stat-big">${plannedMin ? plannedMin + ' min' : '--'}</span>
            <span class="today-stat-big">${fmtDistanceKm(p.distance_km)}</span>
            <span class="today-stat-big">${tss} <span class="today-stat-unit">TSS</span></span>
          </div>`;
        card.style.cursor = 'pointer';
        card.addEventListener('click', () => openWorkoutModal({ source: 'planned', data: p, planned: p, pair: null }));
        feed.appendChild(card);
      });
    }

    function renderTodayFeed() {
      const today = todayKey();
      const yesterday = yesterdayKey();
      const feed = document.getElementById('todayFeed');
      feed.innerHTML = '';

      const metricsItems = calendarItems.filter(i => i.kind === 'metrics' && i.date === today);
      const doneToday = activities.filter(a => dateKeyFromDate(new Date(a.start_date_local)) === today);
      const plannedToday = calendarItems.filter(i => i.kind === 'workout' && i.date === today)
        .sort((a, b) => (a.date > b.date ? 1 : -1));

      // Yesterday's missed planned workouts (no pair, no manual completion)
      const missedYesterday = calendarItems.filter(i => {
        if (i.kind !== 'workout' || i.date !== yesterday) return false;
        const pair = pairForPlanned(String(i.id));
        if (pair) return false;
        if (Number(i.completed_duration_min || 0) > 0) return false;
        return true;
      });

      // Metrics cards
      metricsItems.forEach(m => {
        const lines = (m.description || '').split('\n').map(l => l.trim()).filter(Boolean);
        const rows = lines.map(l => { const [k, ...v] = l.split(':'); return { k: k.trim(), v: v.join(':').trim() }; });
        const card = document.createElement('div');
        card.className = 'today-card today-card-metrics';
        card.innerHTML = `
          <div class="today-card-sport">
            <span>&#9883;</span> ${m.title || 'Metrics'}
          </div>
          ${rows.length ? `<table class="metrics-table">
            ${rows.map(r => `<tr><td class="metrics-key">${r.k}:</td><td class="metrics-val">${r.v}</td></tr>`).join('')}
          </table>` : `<p class="meta" style="padding:8px 16px;">${m.description || ''}</p>`}`;
        card.addEventListener('click', () => openDetailModal(m));
        feed.appendChild(card);
      });

      // Completed activity cards (Strava/imported)
      doneToday.forEach(a => {
        const sport = String(a.type || a.sport_key || 'Workout');
        const tss = Math.round(activityToTss(a));
        const actualMin = Number(a.moving_time || 0) / 60;
        const sportKey = activitySportKey(a);
        const iconSrc = ICON_ASSETS[sportKey] || ICON_ASSETS.other;

        // Find matching planned workout to determine color
        const matchedPair = pairs.find(pr => String(pr.strava_id) === String(a.id));
        const matchedPlanned = matchedPair ? calendarItems.find(i => String(i.id) === String(matchedPair.planned_id)) : null;
        const plannedMin = matchedPlanned ? Number(matchedPlanned.duration_min || 0) : 0;
        const colorClass = completionColorClass(actualMin, plannedMin);

        const card = document.createElement('div');
        card.className = `today-card today-card-completed ${colorClass}`;
        card.innerHTML = `
          <div class="today-card-sport">${sport}</div>
          <div class="today-card-stats">
            <img class="today-card-icon" src="${iconSrc}" alt="${sport}" />
            <span class="today-stat-big">${hms(Number(a.moving_time || 0))}${matchedPlanned ? '&#10003;' : ''}</span>
            <span class="today-stat-big">${fmtDistanceMeters(a.distance)}</span>
            <span class="today-stat-big">${tss} <span class="today-stat-unit">TSS</span></span>
          </div>`;
        card.style.cursor = 'pointer';
        card.addEventListener('click', () => openWorkoutModal({ source: 'strava', data: a }));
        card.addEventListener('contextmenu', ev => showItemMenu(ev, { source: 'strava', data: a }));
        feed.appendChild(card);
      });

      // Planned workout cards
      plannedToday.forEach(p => {
        const pair = pairForPlanned(String(p.id));
        const pairedA = pair ? activities.find(a => String(a.id) === String(pair.strava_id)) : null;
        const sportKey = workoutTypeSportKey(p.workout_type);
        const iconSrc = ICON_ASSETS[sportKey] || ICON_ASSETS.other;
        const plannedMin = Number(p.duration_min || 0);

        const completedDur = Number(p.completed_duration_min || 0);
        const manuallyCompleted = !pairedA && completedDur > 0;

        const card = document.createElement('div');

        if (manuallyCompleted) {
          const completedTss = Number(p.completed_tss || 0) || Math.round(estimateTss(completedDur, 0.75));
          const completedDist = fmtDistanceKm(p.completed_distance_km);
          const colorClass = completionColorClass(completedDur, plannedMin);
          card.className = `today-card today-card-completed ${colorClass}`;
          card.innerHTML = `
            <div class="today-card-sport">
              ${p.workout_type || 'Workout'}
            </div>
            <div class="today-card-stats">
              <img class="today-card-icon" src="${iconSrc}" alt="${p.workout_type}" />
              <span class="today-stat-big">${hms(completedDur * 60)}&#10003;</span>
              <span class="today-stat-big">${completedDist}</span>
              <span class="today-stat-big">${completedTss} <span class="today-stat-unit">TSS</span></span>
            </div>`;
        } else {
          const tss = itemToTss(p);
          // Paired with Strava activity — show completed stats with color
          if (pairedA) {
            const actualMin = Number(pairedA.moving_time || 0) / 60;
            const colorClass = completionColorClass(actualMin, plannedMin);
            const tssA = Math.round(activityToTss(pairedA));
            card.className = `today-card today-card-completed ${colorClass}`;
            card.innerHTML = `
              <div class="today-card-sport">${p.workout_type || 'Workout'}</div>
              <div class="today-card-stats">
                <img class="today-card-icon" src="${iconSrc}" alt="${p.workout_type}" />
                <span class="today-stat-big">${hms(Number(pairedA.moving_time || 0))}&#10003;</span>
                <span class="today-stat-big">${fmtDistanceMeters(pairedA.distance)}</span>
                <span class="today-stat-big">${tssA} <span class="today-stat-unit">TSS</span></span>
              </div>`;
          } else {
            card.className = 'today-card today-card-planned';
            card.innerHTML = `
              <div class="today-card-sport today-card-sport-planned">
                ${p.workout_type || 'Workout'}
                <span class="badge planned" style="margin-left:auto;">Planned</span>
              </div>
              <div class="today-card-stats">
                <img class="today-card-icon" src="${iconSrc}" alt="${p.workout_type}" />
                <span class="today-stat-big">${plannedMin ? plannedMin + ' min' : '--'}</span>
                <span class="today-stat-big">${fmtDistanceKm(p.distance_km)}</span>
                <span class="today-stat-big">${tss} <span class="today-stat-unit">TSS</span></span>
              </div>`;
          }
        }

        card.style.cursor = 'pointer';
        card.addEventListener('click', () => openWorkoutModal({ source: pairedA ? 'strava' : 'planned', data: pairedA || p, planned: p, pair }));
        card.addEventListener('contextmenu', ev => showItemMenu(ev, { source: 'planned', data: p }));
        feed.appendChild(card);
      });

      // Yesterday's missed planned workouts — shown as red
      missedYesterday.forEach(p => {
        const sportKey = workoutTypeSportKey(p.workout_type);
        const iconSrc = ICON_ASSETS[sportKey] || ICON_ASSETS.other;
        const plannedMin = Number(p.duration_min || 0);
        const tss = itemToTss(p);
        const card = document.createElement('div');
        card.className = 'today-card today-card-red';
        card.innerHTML = `
          <div class="today-card-sport">
            ${p.workout_type || 'Workout'}
            <span class="badge" style="margin-left:auto;background:#fde9e9;color:#c0392b;">Missed</span>
          </div>
          <div class="today-card-stats">
            <img class="today-card-icon" src="${iconSrc}" alt="${p.workout_type}" />
            <span class="today-stat-big">${plannedMin ? plannedMin + ' min' : '--'}</span>
            <span class="today-stat-big">${fmtDistanceKm(p.distance_km)}</span>
            <span class="today-stat-big">${tss} <span class="today-stat-unit">TSS</span></span>
          </div>`;
        card.style.cursor = 'pointer';
        card.addEventListener('click', () => openWorkoutModal({ source: 'planned', data: p, planned: p, pair: null }));
        card.addEventListener('contextmenu', ev => showItemMenu(ev, { source: 'planned', data: p }));
        feed.appendChild(card);
      });

      if (!metricsItems.length && !doneToday.length && !plannedToday.length && !missedYesterday.length) {
        feed.innerHTML = '<p class="meta" style="padding:16px 18px;">Nothing logged for today yet.</p>';
      }

      renderTrainingZones(doneToday, plannedToday);
    }

    function renderHome() {
      renderTodayFeed();
      renderTomorrowFeed();
      renderEvents();
      renderGoals();
      renderPerformanceMetrics();
    }

    function renderCalendar(options = {}) {
      const { preserveScroll = true, anchorDate = '', jumpToDate = '' } = options;
      if (preserveScroll) rememberCalendarPosition();
      const dayMap = buildDayAggregateMap();
      const plannedById = new Map(calendarItems.filter(i => i.kind === 'workout').map(i => [String(i.id), i]));
      const stravaById = new Map(activities.map(a => [String(a.id), a]));
      const pairByPlannedId = new Map(pairs.map(p => [String(p.planned_id), p]));
      const pairByStravaId = new Map(pairs.map(p => [String(p.strava_id), p]));
      const wrap = getCalendarScrollContainer();
      const restoreScrollTop = preserveScroll ? Number(calendarState.scrollTop || 0) : 0;
      wrap.innerHTML = '';
      const grid = document.createElement('section');
      grid.className = 'month';

      const baseDate = parseDateKey(anchorDate || calendarState.anchorDate || todayKey());
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
        const rowEnd = new Date(weekStart);
        rowEnd.setDate(weekStart.getDate() + 6);
        const t = parseDateKey(todayKey());
        if (t >= weekStart && t <= rowEnd) row.classList.add('week-current');

        for (let col = 0; col < 7; col += 1) {
          const dayDate = new Date(weekStart);
          dayDate.setDate(weekStart.getDate() + col);
          const key = dateKeyFromDate(dayDate);
          weekDateKeys.push(key);

          const cell = document.createElement('div');
          cell.className = 'day';
          cell.dataset.date = key;
          if (key === todayKey()) cell.classList.add('today');
          if (dayDate.getMonth() !== baseDate.getMonth()) cell.style.opacity = '0.92';

          const num = document.createElement('span');
          num.className = 'd-num';
          num.textContent = key === todayKey() ? `Today ${dayDate.getDate()}` : String(dayDate.getDate());
          const dayHead = document.createElement('div');
          dayHead.className = 'day-head';
          dayHead.appendChild(num);
          cell.appendChild(dayHead);

          // Cell-level drag fallback: catches drops anywhere in the day cell
          cell.addEventListener('dragover', (ev) => ev.preventDefault());
          cell.addEventListener('drop', (ev) => {
            ev.preventDefault();
            const dragData = currentDragData;
            if (!dragData || dragData.source !== 'strava') return;
            // Remove any lingering drop-target highlights
            cell.querySelectorAll('.drop-target').forEach(el => el.classList.remove('drop-target'));
            // Find the single unpaired planned workout for this day
            const unpairedPlanned = calendarItems.filter(i =>
              i.kind === 'workout' && i.date === key && !pairByPlannedId.get(String(i.id))
            );
            if (unpairedPlanned.length === 1) {
              confirmAndPair(String(unpairedPlanned[0].id), String(dragData.id));
            }
          });

          const entries = dayMap[key] || { done: [], items: [] };
          const shownCompleted = new Set();
          const cardsToShow = [];
          // Group all goals for the day into a single card
          const goalItems = entries.items.filter(i => i.kind === 'goal');
          if (goalItems.length > 0) cardsToShow.push({ kind: 'goal-group', items: goalItems });
          entries.items.forEach((item) => {
            if (item.kind === 'goal') return; // already handled above
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
            if (entry.kind === 'goal-group') {
              const items = entry.items;
              const card = document.createElement('div');
              card.className = 'work-card goal';
              const bullets = items.map(g =>
                `<li class="wc-goal-item${g.completed ? ' wc-goal-done' : ''}">${g.title || 'Goal'}</li>`
              ).join('');
              card.innerHTML = `
                <button class="card-menu-btn" type="button">&#8942;</button>
                <div class="wc-kind-head wc-kind-goal"><span class="wc-kind-icon">&#9745;</span> Goals</div>
                <ul class="wc-goal-list">${bullets}</ul>`;
              card.addEventListener('click', (ev) => {
                ev.stopPropagation();
                selectedKind = 'goal'; selectedDate = items[0].date;
                openDetailModal(items[0]);
              });
              card.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'planned', data: items[0] }));
              card.querySelector('.card-menu-btn').addEventListener('click', (ev) => showItemMenu(ev, { source: 'planned', data: items[0] }));
              cell.appendChild(card);
              return;
            }
            if (entry.kind === 'other') {
              const item = entry.item;
              const card = document.createElement('div');
              card.className = `work-card ${item.kind}`;

              if (item.kind === 'event') {
                const d = new Date(item.date + 'T00:00:00');
                const month = d.toLocaleString('en', { month: 'short' }).toUpperCase();
                const day = d.getDate();
                const daysUntil = Math.round((d - new Date(todayKey() + 'T00:00:00')) / 86400000);
                const countdown = daysUntil > 0 ? `${daysUntil} DAY${daysUntil !== 1 ? 'S' : ''} UNTIL EVENT`
                  : daysUntil === 0 ? 'EVENT TODAY'
                  : `${Math.abs(daysUntil)} DAY${Math.abs(daysUntil) !== 1 ? 'S' : ''} AGO`;
                const priorityBadge = item.priority && item.priority !== 'C'
                  ? `<span class="wc-priority-badge wc-priority-${item.priority}">${item.priority}</span>` : '';
                card.innerHTML = `
                  <button class="card-menu-btn" type="button">&#8942;</button>
                  <div class="wc-event-layout">
                    <div class="wc-event-badge"><span class="wc-event-month">${month}</span><span class="wc-event-day">${day}</span></div>
                    <div class="wc-event-info">
                      <p class="wc-event-countdown">${countdown}</p>
                      <p class="wc-event-name">${item.title || 'Event'}${priorityBadge}</p>
                    </div>
                  </div>`;
              } else if (item.kind === 'metrics') {
                const lines = (item.description || '').split('\n').map(l => l.trim()).filter(Boolean);
                const preview = lines[0] || item.title || 'Metrics';
                const more = lines.length > 1 ? `<p class="wc-more">${lines.length - 1} more…</p>` : '';
                card.innerHTML = `
                  <button class="card-menu-btn" type="button">&#8942;</button>
                  <div class="wc-kind-head wc-kind-metrics"><span class="wc-kind-icon">&#9883;</span> ${item.title || 'Metrics'}</div>
                  <p class="wc-meta wc-metrics-preview">${preview}</p>${more}`;
              } else if (item.kind === 'note') {
                const preview = (item.description || item.title || '').substring(0, 60);
                card.innerHTML = `
                  <button class="card-menu-btn" type="button">&#8942;</button>
                  <div class="wc-kind-head wc-kind-note"><span class="wc-kind-icon">&#9998;</span> Note</div>
                  <p class="wc-meta">${preview}${preview.length === 60 ? '…' : ''}</p>`;
              } else {
                card.innerHTML = `
                  <button class="card-menu-btn" type="button">&#8942;</button>
                  <p class="wc-title">${item.title || item.kind}</p>`;
              }

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
              const durMin = completed
                ? completedDurationMinValue(completed)
                : Number(item.duration_min || 0);
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
              card.addEventListener('dragstart', (ev) => {
                currentDragData = { source: 'planned', id: String(item.id) };
                ev.dataTransfer.setData('text/plain', JSON.stringify(currentDragData));
                card.classList.add('dragging-active');
              });
              card.addEventListener('dragend', () => { currentDragData = null; card.classList.remove('dragging-active'); });
              card.addEventListener('dragover', (ev) => ev.preventDefault());
              card.addEventListener('dragenter', (ev) => {
                ev.preventDefault();
                if (!completed) card.classList.add('drop-target');
              });
              card.addEventListener('dragleave', (ev) => {
                if (!card.contains(ev.relatedTarget)) card.classList.remove('drop-target');
              });
              card.addEventListener('drop', (ev) => {
                ev.preventDefault();
                ev.stopPropagation();
                card.classList.remove('drop-target');
                const dragData = currentDragData;
                if (dragData && dragData.source === 'strava') confirmAndPair(String(item.id), String(dragData.id));
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
            const hasSelfPlanned = !pairedPlanned && (
              Number(a.duration_min || 0) > 0
              || Number(a.distance_m || 0) > 0
              || Number(a.distance_km || 0) > 0
              || Number(a.planned_tss || 0) > 0
              || Number(a.planned_if || 0) > 0
            );
            const selfPlanned = hasSelfPlanned ? {
              kind: 'workout',
              duration_min: Number(a.duration_min || 0),
              distance_m: Number(a.distance_m || (Number(a.distance_km || 0) * 1000)),
              distance_km: Number(a.distance_km || 0),
              planned_tss: Number(a.planned_tss || 0),
              planned_if: Number(a.planned_if || 0),
            } : null;
            const compStat = pairedPlanned
              ? complianceStatus(pairedPlanned, a, key).cls
              : selfPlanned
                ? complianceStatus(selfPlanned, a, key).cls
                : 'unplanned';
            const unitForCard = String((pairedPlanned && pairedPlanned.distance_unit) || a.distance_unit || distanceUnit || 'km');
            const card = document.createElement('div');
            card.className = `work-card ${compStat}`;
            const feedCount = commentCount(a);
            const feedback = `${feelEmoji(a.feel)} ${Number(a.rpe || 0) > 0 ? a.rpe : ''}`.trim();
            const commentsText = feedCount > 0 ? `💬 x${feedCount}` : '💬';
            const cDur = completedDurationMinValue(a);
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
            card.addEventListener('dragstart', (ev) => {
              currentDragData = { source: 'strava', id: String(a.id) };
              ev.dataTransfer.setData('text/plain', JSON.stringify(currentDragData));
              card.classList.add('dragging-active');
            });
            card.addEventListener('dragend', () => { currentDragData = null; card.classList.remove('dragging-active'); });
            card.addEventListener('dragover', (ev) => ev.preventDefault());
            card.addEventListener('dragenter', (ev) => {
              ev.preventDefault();
              card.classList.add('drop-target');
            });
            card.addEventListener('dragleave', (ev) => {
              if (!card.contains(ev.relatedTarget)) card.classList.remove('drop-target');
            });
            card.addEventListener('drop', (ev) => {
              ev.preventDefault();
              ev.stopPropagation();
              card.classList.remove('drop-target');
              const dragData = currentDragData;
              if (dragData && dragData.source === 'planned') {
                const draggedPlanned = calendarItems.find((i) => String(i.id) === String(dragData.id));
                if (compStat === 'unplanned' && hasPlannedAndCompletedContent(draggedPlanned)) return;
                confirmAndPair(String(dragData.id), String(a.id));
              }
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
        weekCard.innerHTML = `<div class="ws-metrics"><div class="ws-label-row"><span class="ws-ctl">Fitness</span><span class="ws-atl">Fatigue</span><span class="ws-tsb">Form</span></div><div class="ws-value-row"><span class="ws-ctl">${week.ctl} CTL</span><span class="ws-atl">${week.atl} ATL</span><span class="ws-tsb">${week.tsb > 0 ? '+' + week.tsb : week.tsb} TSB</span></div></div><div class="ws-row"><span>Total Duration</span><strong>${week.durationLabel}</strong></div><div class="ws-row"><span>Total TSS</span><strong>${week.tss}</strong></div>`;
        row.appendChild(weekCard);
        grid.appendChild(row);
        weekRows.push(row);
      }

      wrap.appendChild(grid);
      bindCalendarScrollSync();
      calendarState.hasRendered = true;

      if (jumpToDate) {
        requestAnimationFrame(() => {
          const target = wrap.querySelector(`.day[data-date="${jumpToDate}"]`);
          if (target) {
            const row = target.closest('.week-row');
            const wrapTop = wrap.getBoundingClientRect().top;
            const baseScroll = wrap.scrollTop;
            const targetTop = target.getBoundingClientRect().top - wrapTop + baseScroll;
            const rowTop = row ? (row.getBoundingClientRect().top - wrapTop + baseScroll) : targetTop;
            wrap.scrollTop = Math.max(0, Math.round(rowTop));
            calendarState.scrollTop = wrap.scrollTop;
          }
          syncCalendarHeaderFromScroll();
        });
        return;
      }

      if (preserveScroll) {
        requestAnimationFrame(() => {
          const restored = restoreCalendarPositionFromAnchor();
          if (!restored) {
            wrap.scrollTop = restoreScrollTop;
            calendarState.scrollTop = wrap.scrollTop;
          }
          syncCalendarHeaderFromScroll();
        });
        return;
      }

      wrap.scrollTop = restoreScrollTop;
      calendarState.scrollTop = wrap.scrollTop;
      syncCalendarHeaderFromScroll();
    }

    function jumpToCurrentMonth() {
      const today = todayKey();
      calendarState.anchorDate = today;
      renderCalendar({ preserveScroll: false, anchorDate: today, jumpToDate: today });
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
      const lthr = appSettings.lthr || {};
      const setLthr = (id, key) => { const el = document.getElementById(id); if (el) el.value = lthr[key] == null ? '' : String(lthr[key]); };
      setLthr('lthrRide', 'ride');
      setLthr('lthrRun', 'run');
      setLthr('lthrRow', 'row');
      setLthr('lthrGlobal', 'global');
      const system = (appSettings.unit_system === 'imperial'
        || ((appSettings.units || {}).distance === 'mi' && (appSettings.units || {}).elevation === 'ft'))
        ? 'imperial'
        : 'metric';
      document.querySelectorAll('#unitSystemToggle .seg-btn').forEach((btn) => {
        btn.classList.toggle('active', btn.dataset.system === system);
      });
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

    async function loadData() {
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

      updateUnitButtons();
      renderHome();
      if (isCalendarActive()) {
        renderCalendar({ preserveScroll: true });
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
      const activeSystemBtn = document.querySelector('#unitSystemToggle .seg-btn.active');
      const unitSystem = activeSystemBtn ? activeSystemBtn.dataset.system : 'metric';
      const payload = {
        unit_system: unitSystem,
        units: {
          distance: unitSystem === 'imperial' ? 'mi' : 'km',
          elevation: unitSystem === 'imperial' ? 'ft' : 'm',
        },
        ftp: {
          ride: read('ftpRide'),
          run: read('ftpRun'),
          row: read('ftpRow'),
          swim: read('ftpSwim'),
          strength: read('ftpStrength'),
          other: read('ftpOther'),
        },
        lthr: {
          ride: read('lthrRide'),
          run: read('lthrRun'),
          row: read('lthrRow'),
          global: read('lthrGlobal'),
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
      await loadData();
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
      await loadData();
    });
    const accountBtn = document.getElementById('accountMenuBtn');
    const accountMenu = document.getElementById('accountMenu');
    if (accountBtn && accountMenu) {
      accountBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        accountMenu.classList.toggle('hidden');
      });
      document.getElementById('accountSettingsBtn').addEventListener('click', async () => {
        accountMenu.classList.add('hidden');
        document.getElementById('settingsModal').classList.add('open');
        syncThemeSettingToggle();
        try {
          const res = await fetch('/strava-status');
          const data = await res.json();
          const dot = document.getElementById('stravaStatusDot');
          const label = document.getElementById('stravaStatusLabel');
          const btn = document.getElementById('stravaConnectBtn');
          if (data.connected) {
            dot.style.background = '#17733e';
            label.textContent = data.athlete_name ? `Connected as ${data.athlete_name}` : 'Connected';
            btn.textContent = 'Reconnect';
          } else {
            dot.style.background = '#cc4b37';
            label.textContent = 'Not connected';
            btn.textContent = 'Connect Strava';
          }
        } catch {
          document.getElementById('stravaStatusLabel').textContent = 'Unable to check status';
        }
      });
    }
    const homeUserRow = document.getElementById('homeUserRow');
    if (homeUserRow) {
      homeUserRow.addEventListener('click', async () => {
        document.getElementById('settingsModal').classList.add('open');
        syncThemeSettingToggle();
        try {
          const res = await fetch('/strava-status');
          const data = await res.json();
          const dot = document.getElementById('stravaStatusDot');
          const label = document.getElementById('stravaStatusLabel');
          const btn = document.getElementById('stravaConnectBtn');
          if (data.connected) {
            dot.style.background = '#17733e';
            label.textContent = data.athlete_name ? `Connected as ${data.athlete_name}` : 'Connected';
            btn.textContent = 'Reconnect';
          } else {
            dot.style.background = '#cc4b37';
            label.textContent = 'Not connected';
            btn.textContent = 'Connect Strava';
          }
        } catch {
          document.getElementById('stravaStatusLabel').textContent = 'Unable to check status';
        }
      });
    }

    const addTodayBtnEl = document.getElementById('addTodayBtn');
    if (addTodayBtnEl) {
      addTodayBtnEl.addEventListener('click', () => openActionModal(todayKey()));
    }

    document.querySelectorAll('#unitSystemToggle .seg-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#unitSystemToggle .seg-btn').forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
      });
    });
    document.querySelectorAll('#dEventPriority .seg-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#dEventPriority .seg-btn').forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
      });
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

    // Detail modal close button (X) with unsaved changes check
    function captureDetailInitialState() {
      detailInitialState = {
        title: document.getElementById('dTitle').value,
        date: document.getElementById('dDate').value,
        description: document.getElementById('dDescriptionOther').value,
        wDescription: document.getElementById('dDescription').value,
      };
    }

    function hasUnsavedDetailChanges() {
      if (!detailInitialState) return false;
      return (
        document.getElementById('dTitle').value !== detailInitialState.title ||
        document.getElementById('dDate').value !== detailInitialState.date ||
        document.getElementById('dDescriptionOther').value !== detailInitialState.description ||
        document.getElementById('dDescription').value !== detailInitialState.wDescription
      );
    }

    async function closeDetailWithUnsavedCheck() {
      if (!hasUnsavedDetailChanges()) {
        closeDetailModal();
        return;
      }
      const choice = await confirmUnsavedClose();
      if (choice === 'discard') closeDetailModal();
    }

    document.getElementById('closeDetail').addEventListener('click', closeDetailWithUnsavedCheck);
    document.getElementById('deleteDetail').addEventListener('click', deleteCurrentDetail);
    document.getElementById('saveDetail').addEventListener('click', () => saveDetail(false));
    document.getElementById('saveCloseDetail').addEventListener('click', () => saveDetail(true));
    document.getElementById('detailModal').addEventListener('click', (event) => {
      if (event.target.id === 'detailModal') closeDetailModal();
    });

    async function closeWorkoutWithUnsavedFlow() {
      if (!window.currentWorkoutPayload) {
        await closeWorkoutModal(false);
        return;
      }
      if (!hasUnsavedWorkoutChanges()) {
        await handleSaveAndClose();
        return;
      }
      const choice = await confirmUnsavedClose();
      if (choice === 'discard') {
        if (analyzeState && analyzeState.pendingDirty && typeof analyzeState.cancelPending === 'function') {
          analyzeState.cancelPending();
        }
        await closeWorkoutModal(true);
        await loadData();
      }
    }

    document.getElementById('closeWorkoutView').addEventListener('click', closeWorkoutWithUnsavedFlow);
    document.getElementById('saveWorkoutView').addEventListener('click', handleSave);
    document.getElementById('saveCloseWorkoutView').addEventListener('click', handleSaveAndClose);
    document.getElementById('deleteWorkoutView').addEventListener('click', () => {
      const payload = window.currentWorkoutPayload;
      if (!payload) return;
      confirmDelete({ onConfirm: async () => {
        const data = payload.data || {};
        if (payload.planned || payload.source === 'planned') {
          const targetId = payload.planned ? payload.planned.id : data.id;
          if (!targetId) { await closeWorkoutModal(false); return; }
          await fetch(`/calendar-items/${targetId}`, { method: 'DELETE' });
        } else if (payload.source === 'strava') {
          await fetch(`/activities/${data.id}`, { method: 'DELETE' });
        }
        await closeWorkoutModal(false);
        await loadData();
      } });
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
        const rpeVal = Number(document.getElementById('wvRpe').value || 0);
        renderTopFeelRpe(btn.dataset.feel, (modalDraft && modalDraft.rpeTouched) ? rpeVal : 0);
      });
    });
    document.getElementById('wvRpe').addEventListener('input', () => {
      if (modalDraft) modalDraft.rpeTouched = true;
      const rpe = document.getElementById('wvRpe');
      rpe.classList.remove('rpe-unset');
      updateRpeLabel();
      renderTopFeelRpe(currentFeel, Number(rpe.value || 0));
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
    document.getElementById('wvCommentsFeed').addEventListener('click', async (ev) => {
      const del = ev.target.closest('.comment-delete');
      if (!del || !modalDraft) return;
      const item = del.closest('.comment-item');
      if (!item) return;
      const idx = Number(item.dataset.index);
      if (!Number.isInteger(idx) || idx < 0) return;
      const ok = await confirmDeleteComment();
      if (!ok) return;
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
        closeWorkoutWithUnsavedFlow();
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
      if (!ev.target.closest('#accountWrap') && accountMenu) {
        accountMenu.classList.add('hidden');
      }
    });
    document.getElementById('calTodayBtn').addEventListener('click', jumpToCurrentMonth);

    // ── Delete confirm modal wiring ──
    document.getElementById('deleteConfirmCancel').addEventListener('click', async () => {
      document.getElementById('deleteConfirmModal').classList.remove('open');
      if (_deleteConfirmResolver) {
        const resolver = _deleteConfirmResolver;
        _deleteConfirmResolver = null;
        await resolver(false);
      }
    });
    document.getElementById('deleteConfirmOk').addEventListener('click', async () => {
      document.getElementById('deleteConfirmModal').classList.remove('open');
      if (_deleteConfirmResolver) {
        const resolver = _deleteConfirmResolver;
        _deleteConfirmResolver = null;
        await resolver(true);
      }
    });
    document.getElementById('deleteConfirmModal').addEventListener('click', async (ev) => {
      if (ev.target.id === 'deleteConfirmModal') {
        document.getElementById('deleteConfirmModal').classList.remove('open');
        if (_deleteConfirmResolver) {
          const resolver = _deleteConfirmResolver;
          _deleteConfirmResolver = null;
          await resolver(false);
        }
      }
    });

    // ── Apply confirm modal wiring ──
    document.getElementById('applyConfirmCancel').addEventListener('click', () => {
      document.getElementById('applyConfirmModal').classList.remove('open');
      if (_applyConfirmResolver) {
        const resolver = _applyConfirmResolver;
        _applyConfirmResolver = null;
        resolver(false);
      }
    });
    document.getElementById('applyConfirmOk').addEventListener('click', () => {
      document.getElementById('applyConfirmModal').classList.remove('open');
      if (_applyConfirmResolver) {
        const resolver = _applyConfirmResolver;
        _applyConfirmResolver = null;
        resolver(true);
      }
    });
    document.getElementById('applyConfirmModal').addEventListener('click', (ev) => {
      if (ev.target.id === 'applyConfirmModal') {
        document.getElementById('applyConfirmModal').classList.remove('open');
        if (_applyConfirmResolver) {
          const resolver = _applyConfirmResolver;
          _applyConfirmResolver = null;
          resolver(false);
        }
      }
    });

    // ── Comment delete confirm modal wiring ──
    document.getElementById('commentDeleteConfirmCancel').addEventListener('click', () => {
      document.getElementById('commentDeleteConfirmModal').classList.remove('open');
      if (_commentDeleteResolver) {
        const resolver = _commentDeleteResolver;
        _commentDeleteResolver = null;
        resolver(false);
      }
    });
    document.getElementById('commentDeleteConfirmOk').addEventListener('click', () => {
      document.getElementById('commentDeleteConfirmModal').classList.remove('open');
      if (_commentDeleteResolver) {
        const resolver = _commentDeleteResolver;
        _commentDeleteResolver = null;
        resolver(true);
      }
    });
    document.getElementById('commentDeleteConfirmModal').addEventListener('click', (ev) => {
      if (ev.target.id === 'commentDeleteConfirmModal') {
        document.getElementById('commentDeleteConfirmModal').classList.remove('open');
        if (_commentDeleteResolver) {
          const resolver = _commentDeleteResolver;
          _commentDeleteResolver = null;
          resolver(false);
        }
      }
    });

    // ── Unsaved-close confirm modal wiring ──
    document.getElementById('unsavedCloseCancel').addEventListener('click', () => {
      document.getElementById('unsavedCloseConfirmModal').classList.remove('open');
      if (_unsavedCloseResolver) {
        const resolver = _unsavedCloseResolver;
        _unsavedCloseResolver = null;
        resolver('cancel');
      }
    });
    document.getElementById('unsavedCloseDiscard').addEventListener('click', () => {
      document.getElementById('unsavedCloseConfirmModal').classList.remove('open');
      if (_unsavedCloseResolver) {
        const resolver = _unsavedCloseResolver;
        _unsavedCloseResolver = null;
        resolver('discard');
      }
    });
    document.getElementById('unsavedCloseConfirmModal').addEventListener('click', (ev) => {
      if (ev.target.id === 'unsavedCloseConfirmModal') {
        document.getElementById('unsavedCloseConfirmModal').classList.remove('open');
        if (_unsavedCloseResolver) {
          const resolver = _unsavedCloseResolver;
          _unsavedCloseResolver = null;
          resolver('cancel');
        }
      }
    });

    // ── Dark mode toggle ──
    function initDarkMode() {
      const savedTheme = localStorage.getItem('theme') || 'light';
      document.documentElement.setAttribute('data-theme', savedTheme);
      // Theme toggle is now in Settings panel (themeSettingToggle)
    }

    function syncThemeSettingToggle() {
      const toggle = document.getElementById('themeSettingToggle');
      if (!toggle) return;
      const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
      toggle.querySelectorAll('.seg-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.themeVal === currentTheme);
      });
    }

    function initThemeSettingToggle() {
      const toggle = document.getElementById('themeSettingToggle');
      if (!toggle) return;
      syncThemeSettingToggle();
      toggle.querySelectorAll('.seg-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const newTheme = btn.dataset.themeVal;
          document.documentElement.setAttribute('data-theme', newTheme);
          localStorage.setItem('theme', newTheme);
          syncThemeSettingToggle();
        });
      });
    }

    initDarkMode();
    initThemeSettingToggle();
    buildTypeGrids();
    bindWidgetToggles();
    applyWidgetPrefs();
    updateUnitButtons();
    setView('home');
    loadData();
