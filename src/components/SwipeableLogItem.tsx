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
  const currentX = useRef(0);
  const isDragging = useRef(false);

  const handleTouchStart = (e: React.TouchEvent) => {
    if (disabled) return;
    startX.current = e.touches[0].clientX;
    isDragging.current = true;
  };

  const handleTouchMove = (e: React.TouchEvent) => {
    if (!isDragging.current || disabled) return;
    currentX.current = e.touches[0].clientX;
    const diff = currentX.current - startX.current;

    // Only allow dragging left
    if (diff < 0) {
      // Prevent default to stop scrolling while swiping horizontally
      if (Math.abs(diff) > 10 && e.cancelable) {
        // We can't preventDefault in React synthetic events easily if passive: true
        // But we rely on touchAction: 'pan-y' to handle this mostly.
      }
      const newX = Math.max(diff, -actionWidth - 20); // Add some elasticity
      controls.set({ x: isRevealed ? -actionWidth + newX : newX });
    } else if (isRevealed && diff > 0) {
      const newX = Math.min(-actionWidth + diff, 0);
      controls.set({ x: newX });
    }
  };

  const handleTouchEnd = () => {
    if (!isDragging.current || disabled) return;
    isDragging.current = false;

    const diff = currentX.current - startX.current;

    // Use a more standard iOS-like spring
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
      className="relative w-full overflow-hidden rounded-2xl bg-black/5"
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
        className="relative z-10 w-full rounded-2xl bg-white"
        style={{ touchAction: "pan-y" }}
      >
        {children}
      </motion.div>
    </div>
  );
};

export default SwipeableLogItem;
