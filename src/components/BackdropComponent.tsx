"use client";

import { ReactNode } from "react";

interface BackdropProps {
  open: boolean;
  children?: ReactNode;
}

export default function BackdropComponent({ open, children }: BackdropProps) {
  if (!open) return null;

  return (
    // <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-40">
    <div className="fixed inset-0 z-40 flex items-center justify-center bg-black/30">
      {children}
    </div>
  );
}
