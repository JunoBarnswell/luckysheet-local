import React, { useRef } from 'react';
import { Button, type ButtonProps } from './Button';

export interface FileButtonProps extends Omit<ButtonProps, 'onChange'> {
  accept?: string;
  onFile: (file: globalThis.File) => void;
}

/** 文件选择按钮:内部封装隐藏的 file input,业务组件无需直接使用原生元素 */
export function FileButton({ accept, onFile, ...buttonProps }: FileButtonProps): React.ReactElement {
  const inputRef = useRef<HTMLInputElement | null>(null);

  return (
    <>
      <input
        ref={inputRef}
        accept={accept}
        className="hidden"
        type="file"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFile(file);
          event.target.value = '';
        }}
      />
      <Button {...buttonProps} onClick={() => inputRef.current?.click()} />
    </>
  );
}
