export function setStatus(target: HTMLElement, title: string, detail: string): void {
  const heading = document.createElement('strong');
  heading.textContent = title;
  const body = document.createElement('span');
  body.textContent = detail;
  target.replaceChildren(heading, body);
}

export function setError(target: HTMLElement, message: string): void {
  const wrapper = document.createElement('div');
  wrapper.className = 'sb-demo__empty';
  const content = document.createElement('div');
  setStatus(content, 'The workbook could not be processed', message);
  wrapper.append(content);
  target.replaceChildren(wrapper);
}
