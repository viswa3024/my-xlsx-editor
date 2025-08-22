"use client";

export default function LoadingComponent() {
  return (
    <div className="flex flex-col items-center justify-center">
      <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin"></div>
      <p className="mt-3 text-white font-medium">Generating...</p>
    </div>
  );
}
