export const throttle = (fun: Function, delay: number) => {
  let timeout: any = null;
  return function (...args: any[]) {
    const context = this;
    if(!timeout) {
      timeout = setTimeout(() => {
        clearTimeout(timeout);
        timeout = null;
        fun.apply(context, args);
      }, delay)
    }
  }
}