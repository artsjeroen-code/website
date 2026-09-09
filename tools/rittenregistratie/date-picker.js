(() => {
  'use strict';

  const TARGET_IDS = ['rideDate', 'vehicleUseFrom', 'vehicleUseTo', 'correctionDate'];
  const MONTHS = ['januari', 'februari', 'maart', 'april', 'mei', 'juni', 'juli', 'augustus', 'september', 'oktober', 'november', 'december'];
  const WEEKDAYS = ['Ma', 'Di', 'Wo', 'Do', 'Vr', 'Za', 'Zo'];
  const enhanced = [];
  let openPicker = null;

  function pad(value) {
    return String(value).padStart(2, '0');
  }

  function parseIso(value) {
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ''));
    if (!match) return null;
    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    if (date.getFullYear() !== Number(match[1]) || date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[3])) return null;
    return date;
  }

  function isoValue(date) {
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  }

  function displayValue(value) {
    const date = parseIso(value);
    return date ? `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()}` : '';
  }

  function closePicker(picker = openPicker) {
    if (!picker) return;
    picker.popup.hidden = true;
    picker.display.setAttribute('aria-expanded', 'false');
    if (openPicker === picker) openPicker = null;
  }

  function dateAllowed(picker, date) {
    const value = isoValue(date);
    return (!picker.input.min || value >= picker.input.min) && (!picker.input.max || value <= picker.input.max);
  }

  function render(picker) {
    const first = new Date(picker.viewYear, picker.viewMonth, 1);
    const startOffset = (first.getDay() + 6) % 7;
    const gridStart = new Date(picker.viewYear, picker.viewMonth, 1 - startOffset);
    const selected = parseIso(picker.input.value);
    const todayIso = isoValue(new Date());

    picker.title.textContent = `${MONTHS[picker.viewMonth]} ${picker.viewYear}`;
    picker.days.replaceChildren();

    for (let index = 0; index < 42; index += 1) {
      const date = new Date(gridStart.getFullYear(), gridStart.getMonth(), gridStart.getDate() + index);
      const value = isoValue(date);
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'calendar-day';
      button.textContent = String(date.getDate());
      button.dataset.outside = String(date.getMonth() !== picker.viewMonth);
      if (value === todayIso) button.dataset.today = 'true';
      if (selected && value === isoValue(selected)) button.dataset.selected = 'true';
      button.disabled = !dateAllowed(picker, date);
      button.setAttribute('aria-label', `${date.getDate()} ${MONTHS[date.getMonth()]} ${date.getFullYear()}`);
      button.addEventListener('click', () => {
        picker.input.value = value;
        picker.display.value = displayValue(value);
        picker.input.dispatchEvent(new Event('input', { bubbles: true }));
        picker.input.dispatchEvent(new Event('change', { bubbles: true }));
        closePicker(picker);
      });
      picker.days.appendChild(button);
    }
  }

  function open(picker) {
    if (openPicker && openPicker !== picker) closePicker(openPicker);
    const selected = parseIso(picker.input.value) || new Date();
    picker.viewYear = selected.getFullYear();
    picker.viewMonth = selected.getMonth();
    render(picker);
    picker.popup.hidden = false;
    picker.display.setAttribute('aria-expanded', 'true');
    openPicker = picker;
  }

  function enhance(input) {
    if (!input || input.dataset.nlDatePicker === 'true') return;
    input.dataset.nlDatePicker = 'true';
    input.classList.add('date-native-input');

    const wrapper = document.createElement('div');
    wrapper.className = 'date-picker-field';
    input.parentNode.insertBefore(wrapper, input);
    wrapper.appendChild(input);

    const display = document.createElement('input');
    display.type = 'text';
    display.className = 'date-display-input';
    display.readOnly = true;
    display.inputMode = 'none';
    display.placeholder = 'DD/MM/YYYY';
    display.value = displayValue(input.value);
    display.setAttribute('aria-label', `${input.closest('label')?.querySelector('span')?.textContent?.trim() || 'Datum'} (DD/MM/YYYY)`);
    display.setAttribute('aria-haspopup', 'dialog');
    display.setAttribute('aria-expanded', 'false');
    wrapper.appendChild(display);

    const popup = document.createElement('div');
    popup.className = 'date-picker-popup';
    popup.hidden = true;
    popup.setAttribute('role', 'dialog');
    popup.setAttribute('aria-label', 'Datum kiezen');

    const header = document.createElement('div');
    header.className = 'calendar-header';
    const previous = document.createElement('button');
    previous.type = 'button';
    previous.className = 'calendar-nav';
    previous.textContent = '‹';
    previous.setAttribute('aria-label', 'Vorige maand');
    const title = document.createElement('strong');
    title.className = 'calendar-title';
    const next = document.createElement('button');
    next.type = 'button';
    next.className = 'calendar-nav';
    next.textContent = '›';
    next.setAttribute('aria-label', 'Volgende maand');
    header.append(previous, title, next);

    const weekdays = document.createElement('div');
    weekdays.className = 'calendar-weekdays';
    WEEKDAYS.forEach((weekday) => {
      const item = document.createElement('span');
      item.textContent = weekday;
      weekdays.appendChild(item);
    });

    const days = document.createElement('div');
    days.className = 'calendar-days';
    popup.append(header, weekdays, days);
    wrapper.appendChild(popup);

    const picker = { input, display, popup, title, days, viewYear: 0, viewMonth: 0 };
    enhanced.push(picker);

    display.addEventListener('click', () => open(picker));
    display.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        open(picker);
      }
    });
    previous.addEventListener('click', () => {
      picker.viewMonth -= 1;
      if (picker.viewMonth < 0) {
        picker.viewMonth = 11;
        picker.viewYear -= 1;
      }
      render(picker);
    });
    next.addEventListener('click', () => {
      picker.viewMonth += 1;
      if (picker.viewMonth > 11) {
        picker.viewMonth = 0;
        picker.viewYear += 1;
      }
      render(picker);
    });
  }

  function syncAll() {
    enhanced.forEach((picker) => {
      const formatted = displayValue(picker.input.value);
      if (picker.display.value !== formatted) picker.display.value = formatted;
    });
  }

  TARGET_IDS.forEach((id) => enhance(document.getElementById(id)));

  document.addEventListener('click', (event) => {
    if (openPicker && !openPicker.popup.parentElement.contains(event.target)) closePicker(openPicker);
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') closePicker();
  });
  window.setInterval(syncAll, 300);
})();
