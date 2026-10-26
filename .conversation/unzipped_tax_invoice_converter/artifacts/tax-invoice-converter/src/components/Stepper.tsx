import { Check } from 'lucide-react';

export function Stepper({ step, labels }: { step: number; labels: string[] }) {
  return <ol className="stepper" aria-label="Progress">
    {labels.map((label, i) => <li key={label} className={`step-item ${step >= i + 1 ? 'is-active' : ''} ${step > i + 1 ? 'is-done' : ''}`}>
      <span className="step-circle">{step > i + 1 ? <Check size={15} aria-hidden="true" /> : i + 1}</span>
      <span className="step-label">{label}</span>
      {i < labels.length - 1 && <span className="step-line" aria-hidden="true" />}
    </li>)}
  </ol>;
}
