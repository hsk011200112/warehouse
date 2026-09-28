import React, { useState, useRef, useEffect } from "react";
import { motion, useAnimation } from "motion/react";
import { Trash2, Edit2 } from "lucide-react";

interface SwipeableLogItemProps {
  children: React.ReactNode;
  onDelete: () => void;
  onEdit: () => void;
  disabled?: boolean;
}

const SwipeableLogItem: React.FC<SwipeableLogItemProps> = ({
  children,
  onDelete,
  onEdit,
  disabled = false,
}) => {
  const controls = useAnimation();
  const [isRevealed, setIsRevealed] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);
  const contentRef = useRef<HTMLDivElement>(null);

  // The width of the action buttons (Edit + Delete)
  const actionWidth = 120; // 60px per button

  const startX = useRef(0);
  const startY = useRef(0);
  const currentX = useRef(0);
  const isDragging = useRef(false);
  const gestureDirection = useRef<'horizontal' | 'vertical' | null>(null);

  const handleTouchStart = (e: React.TouchEvent) => {
    if (disabled) return;
    startX.current = e.touches[0].clientX;
    startY.current = e.touches[0].clientY;
    currentX.current = e.touches[0].clientX;
    isDragging.current = true;
    gestureDirection.current = null;
  };

  const handleTouchMove = (e: React.TouchEvent) => {
    if (!isDragging.current || disabled) return;
    const clientX = e.touches[0].clientX;
    const clientY = e.touches[0].clientY;
    currentX.current = clientX;
    
    const diffX = clientX - startX.current;
    const diffY = clientY - startY.current;

    // Detect direction on initial movement threshold (8px)
    if (gestureDirection.current === null) {
      if (Math.abs(diffX) > 8 || Math.abs(diffY) > 8) {
        if (Math.abs(diffX) > Math.abs(diffY)) {
          gestureDirection.current = 'horizontal';
        } else {
          // Detected vertical scrolling: disengage completely so page can scroll freely!
          gestureDirection.current = 'vertical';
          isDragging.current = false;
          return;
        }
      } else {
        return;
      }
    }

    if (gestureDirection.current !== 'horizontal') return;

    // Only allow dragging left (or right to close if already revealed)
    if (diffX < 0) {
      const newX = Math.max(diffX, -actionWidth - 20); // Add slight elasticity
      controls.set({ x: isRevealed ? -actionWidth + newX : newX });
    } else if (isRevealed && diffX > 0) {
      const newX = Math.min(-actionWidth + diffX, 0);
      controls.set({ x: newX });
    }
  };

  const handleTouchEnd = () => {
    if (!isDragging.current || disabled || gestureDirection.current !== 'horizontal') {
      isDragging.current = false;
      gestureDirection.current = null;
      return;
    }
    isDragging.current = false;
    gestureDirection.current = null;

    const diff = currentX.current - startX.current;
    const springConfig = { type: "spring", stiffness: 450, damping: 35 } as const;

    if (!isRevealed && diff < -actionWidth / 2) {
      controls.start({
        x: -actionWidth,
        transition: springConfig,
      });
      setIsRevealed(true);
    } else if (isRevealed && diff > actionWidth / 2) {
      controls.start({
        x: 0,
        transition: springConfig,
      });
      setIsRevealed(false);
    } else {
      // Revert with bounce
      controls.start({
        x: isRevealed ? -actionWidth : 0,
        transition: springConfig,
      });
    }
  };

  // Close when clicking outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent | TouchEvent) => {
      if (
        isRevealed &&
        containerRef.current &&
        !containerRef.current.contains(e.target as Node)
      ) {
        controls.start({
          x: 0,
          transition: { type: "spring", stiffness: 450, damping: 35 } as const,
        });
        setIsRevealed(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    document.addEventListener("touchstart", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
      document.removeEventListener("touchstart", handleClickOutside);
    };
  }, [isRevealed, controls]);

  return (
    <div
      className="relative w-full overflow-hidden rounded-2xl bg-black/5 select-none"
      ref={containerRef}
    >
      {/* Background Actions - iOS Style */}
      <div className="absolute inset-y-0 right-0 flex items-center justify-end w-full">
        <div className="flex h-full items-stretch">
          <button
            onClick={(e) => {
              e.stopPropagation();
              controls.start({ x: 0 });
              setIsRevealed(false);
              onEdit();
            }}
            className="flex items-center justify-center w-[60px] h-full bg-apple-blue text-white active:scale-95 transition-transform"
          >
            <Edit2 size={20} className="drop-shadow-sm" />
          </button>
          <button
            onClick={(e) => {
              e.stopPropagation();
              controls.start({ x: 0 });
              setIsRevealed(false);
              onDelete();
            }}
            className="flex items-center justify-center w-[60px] h-full bg-red-500 text-white active:scale-95 transition-transform"
          >
            <Trash2 size={20} className="drop-shadow-sm" />
          </button>
        </div>
      </div>

      {/* Foreground Content */}
      <motion.div
        ref={contentRef}
        animate={controls}
        onTouchStart={handleTouchStart}
        onTouchMove={handleTouchMove}
        onTouchEnd={handleTouchEnd}
        onTouchCancel={handleTouchEnd}
        className="relative z-10 w-full rounded-2xl bg-white"
        style={{ touchAction: "pan-y" }}
      >
        {children}
      </motion.div>
    </div>
  );
};

export default SwipeableLogItem;
