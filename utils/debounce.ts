export function debounce (fun: Function, delay: number, immediate: boolean) {
  let timeout: any;
  return function (...args: any[]) {
    const context = this;
    timeout && clearTimeout(timeout);
    timeout = setTimeout(() => {
      timeout = null;
      if (!immediate) fun.apply(context, args);
    }, delay);
  }
}