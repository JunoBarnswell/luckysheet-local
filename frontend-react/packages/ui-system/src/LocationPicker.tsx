import { Select, type SelectProps } from './Select';

export interface LocationOption {
  id: string;
  label: string;
  disabled?: boolean;
}

export interface LocationPickerProps extends Omit<SelectProps, 'children' | 'options' | 'placeholder'> {
  options: readonly LocationOption[];
  placeholder?: string;
}

export function LocationPicker({ options, placeholder, ...props }: LocationPickerProps) {
  return <Select {...props} options={options.map((option) => ({ value: option.id, label: option.label, disabled: option.disabled }))} placeholder={placeholder} />;
}
