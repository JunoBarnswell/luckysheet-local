import type { InputHTMLAttributes } from 'react';
import { cn } from './cn';
import { Inline, Text } from './layout';

export interface RadioOptionProps extends Omit<InputHTMLAttributes<HTMLInputElement>, 'type'> {
  label: string;
}

/** Shared semantic radio control; business components never render native inputs directly. */
export function RadioOption({ className, label, ...props }: RadioOptionProps) {
  return (
    <Text as="label" size="sm" className={cn('cursor-pointer', className)}>
      <Inline gap="xs"><input {...props} type="radio" className="h-4 w-4 accent-[#217346]" /><Text size="sm">{label}</Text></Inline>
    </Text>
  );
}
