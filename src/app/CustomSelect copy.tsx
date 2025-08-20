"use client";

import React from "react";

type CustomSelectProps = {
  value: string;
  options: string[];
  onChange: (value: string) => void;
  className?: string;
};

export default function CustomSelect({
  value,
  options,
  onChange,
  className = "",
}: CustomSelectProps) {
  return (
    <select
      className={`w-full p-2 border border-gray-300 rounded-md bg-white text-gray-700 shadow-sm 
        focus:outline-none focus:ring-0 transition cursor-pointer ${className}`}
      value={value}
      onChange={(e) => onChange(e.target.value)}
    >
      {options.map((opt, idx) => (
        <option key={idx} value={opt}>
          {opt}
        </option>
      ))}
    </select>
  );
}
