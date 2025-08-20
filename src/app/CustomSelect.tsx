'use client';

import { ChevronDown } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';

type Option = {
  key: string;
  label: string;
};

type SelectProps = {
  label?: string;
  options: Option[];
  value: string;
  onChange: (key: string) => void;
  className?: string;
};

export default function CustomSelect({
  label,
  options,
  value,
  onChange,
  className = "",
}: SelectProps) {
  const [isOpen, setIsOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  const selected = options.find((opt) => opt.key === value);

  // 🔹 Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (ref.current && !ref.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div className={`relative w-full ${className}`} ref={ref}>
      {label && <label className="block text-gray-700 font-medium mb-1">{label}</label>}

      <button
        type="button"
        onClick={() => setIsOpen((prev) => !prev)}
        className="w-full border border-gray-300 rounded px-3 py-2 text-left flex justify-between items-center bg-white shadow-sm"
      >
        <span className={selected ? "text-gray-800" : "text-gray-400"}>
          {selected?.label || 'Select...'}
        </span>
        <ChevronDown className="h-4 w-4 text-gray-500" />
      </button>

      {isOpen && (
        <ul className="absolute z-10 mt-1 w-full bg-white border border-gray-200 rounded shadow-md max-h-60 overflow-y-auto">
          {options.map((option) => (
            <li
              key={option.key}
              onClick={() => {
                onChange(option.key);
                setIsOpen(false);
              }}
              className={`px-4 py-2 cursor-pointer hover:bg-blue-100 transition ${
                option.key === value ? 'bg-blue-50 font-semibold text-blue-600' : ''
              }`}
            >
              {option.label}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
