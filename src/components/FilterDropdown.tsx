'use client';

import { useState } from 'react';
import { Popover, PopoverContent, PopoverTrigger } from '@/components/ui/popover';
import { Badge } from '@/components/ui/badge';
import { Checkbox } from '@/components/ui/checkbox';
import { Switch } from '@/components/ui/switch';
import { Button } from '@/components/ui/button';
import { Filter, ChevronDown, X } from 'lucide-react';
import { cn } from '@/lib/utils';

interface FilterDropdownProps {
  label: string;
  options: { value: string; label: string }[];
  selected: string[];
  onChange: (selected: string[]) => void;
  maxDisplay?: number;
}

export function FilterDropdown({ label, options, selected, onChange, maxDisplay = 2 }: FilterDropdownProps) {
  const [open, setOpen] = useState(false);
  const [multiple, setMultiple] = useState(false);

  const handleToggle = (value: string) => {
    if (multiple) {
      if (selected.includes(value)) {
        onChange(selected.filter(v => v !== value));
      } else {
        onChange([...selected, value]);
      }
    } else {
      if (selected.includes(value)) {
        onChange([]);
      } else {
        onChange([value]);
      }
      setOpen(false);
    }
  };

  const handleClear = (e: React.MouseEvent) => {
    e.stopPropagation();
    onChange([]);
  };

  const handleSelectAll = () => {
    if (selected.length === options.length) {
      onChange([]);
    } else {
      onChange(options.map(o => o.value));
    }
  };

  const displayLabel = () => {
    if (selected.length === 0) return label;
    if (selected.length === 1) {
      const found = options.find(o => o.value === selected[0]);
      return found?.label || selected[0];
    }
    return `${selected.length} selecionados`;
  };

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button
          variant={selected.length > 0 ? 'default' : 'outline'}
          className={cn(
            'h-9 gap-1.5 text-sm font-normal',
            selected.length > 0 && 'bg-slate-600 hover:bg-slate-700'
          )}
        >
          <Filter className="h-3.5 w-3.5 shrink-0" />
          <span className="truncate max-w-[160px]">{displayLabel()}</span>
          {selected.length > 0 && (
            <Badge variant="secondary" className="ml-0.5 h-5 min-w-[20px] px-1.5 text-[10px]">
              {selected.length}
            </Badge>
          )}
          <ChevronDown className="h-3.5 w-3.5 opacity-50 shrink-0" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-64 p-0" align="start">
        {/* Header */}
        <div className="flex items-center justify-between px-3 py-2 border-b">
          <span className="text-xs font-medium text-slate-500">{label}</span>
          <div className="flex items-center gap-2">
            {selected.length > 0 && (
              <Button variant="ghost" size="sm" className="h-6 px-2 text-xs text-slate-500" onClick={handleClear}>
                <X className="h-3 w-3 mr-0.5" /> Limpar
              </Button>
            )}
          </div>
        </div>

        {/* Multiple toggle */}
        <div className="flex items-center justify-between px-3 py-2 bg-slate-50/50 border-b">
          <label htmlFor={`multi-${label}`} className="text-xs text-slate-600 cursor-pointer flex items-center gap-1.5">
            Múltiplo
          </label>
          <Switch
            id={`multi-${label}`}
            checked={multiple}
            onCheckedChange={setMultiple}
            className="scale-75"
          />
        </div>

        {/* Select all (only in multiple mode) */}
        {multiple && options.length > 1 && (
          <button
            className="w-full text-left px-3 py-1.5 text-xs text-slate-500 hover:bg-slate-100 border-b transition-colors"
            onClick={handleSelectAll}
          >
            {selected.length === options.length ? 'Desmarcar todos' : 'Selecionar todos'}
          </button>
        )}

        {/* Options */}
        <div className="max-h-64 overflow-y-auto py-1">
          {options.map(option => {
            const isSelected = selected.includes(option.value);
            return (
              <button
                key={option.value}
                className={cn(
                  'w-full flex items-center gap-2 px-3 py-1.5 text-sm hover:bg-slate-100 transition-colors text-left',
                  isSelected && 'bg-slate-50'
                )}
                onClick={() => handleToggle(option.value)}
              >
                {multiple ? (
                  <Checkbox checked={isSelected} className="pointer-events-none" />
                ) : (
                  <div className={cn(
                    'h-4 w-4 rounded-full border-2 shrink-0 flex items-center justify-center',
                    isSelected ? 'border-slate-600 bg-slate-600' : 'border-slate-300'
                  )}>
                    {isSelected && <div className="h-1.5 w-1.5 rounded-full bg-white" />}
                  </div>
                )}
                <span className={cn('truncate', isSelected && 'font-medium text-slate-800')}>
                  {option.label}
                </span>
                {isSelected && !multiple && (
                  <X className="h-3 w-3 ml-auto text-slate-400 shrink-0" />
                )}
              </button>
            );
          })}
          {options.length === 0 && (
            <div className="px-3 py-4 text-center text-sm text-slate-400">Sem opções</div>
          )}
        </div>
      </PopoverContent>
    </Popover>
  );
}
