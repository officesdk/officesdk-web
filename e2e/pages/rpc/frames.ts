/**
 * 创建一个 iframe 用于模拟服务端，
 * 并将对应的路径注入到 iframe 的 script 标签中
 */
export function createServerFrame(content: HTMLElement, iframe?: HTMLIFrameElement): HTMLIFrameElement {
  if (!iframe) {
    iframe = document.createElement('iframe');
  }

  iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin');
  iframe.src = '/template.html';

  iframe.classList.add('frame');

  content.appendChild(iframe);

  iframe.onload = () => {
    // 将对应的路径注入到 iframe 的 script 标签中
    const script = document.createElement('script');
    script.src = '/server.bundle.js';
    iframe.contentWindow?.document.body.appendChild(script);
  };

  return iframe;
}

export function createContainerFrame(content: HTMLElement): HTMLDivElement {
  const container = document.createElement('div');
  container.classList.add('frame');
  content.appendChild(container);

  return container;
}
