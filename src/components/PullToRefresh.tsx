import React, { useState, useRef, useEffect, useCallback } from 'react';
import { motion } from 'motion/react';
import { RefreshCw } from 'lucide-react';

interface PullToRefreshProps {
  onRefresh: () => Promise<void>;
  children: React.ReactNode;
}

const PullToRefresh: React.FC<PullToRefreshProps> = ({ onRefresh, children }) => {
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [pullProgress, setPullProgress] = useState(0);
  const containerRef = useRef<HTMLDivElement>(null);
  const contentRef = useRef<HTMLDivElement>(null);

  const startX = useRef(0);
  const startY = useRef(0);
  const currentY = useRef(0);
  const isDragging = useRef(false);
  const isScrollingTrapped = useRef(false);
  const PULL_THRESHOLD = 80;

  const handleTouchStart = useCallback((e: TouchEvent) => {
    if (containerRef.current?.scrollTop === 0 && !isRefreshing) {
      startX.current = e.touches[0].clientX;
      startY.current = e.touches[0].clientY;
      isDragging.current = true;
      isScrollingTrapped.current = false;
    }
  }, [isRefreshing]);

  const handleTouchMove = useCallback((e: TouchEvent) => {
    if (!isDragging.current || isRefreshing) return;

    const currentX = e.touches[0].clientX;
    currentY.current = e.touches[0].clientY;
    
    const diffX = Math.abs(currentX - startX.current);
    const diffY = Math.abs(currentY.current - startY.current);
    const actualDiffY = currentY.current - startY.current;

    // Fast bail out of pull-to-refresh if horizontal scrolling is detected
    if (diffX > diffY && diffX > 5) {
      isScrollingTrapped.current = true;
    }

    if (isScrollingTrapped.current) return;

    if (actualDiffY > 0 && containerRef.current?.scrollTop === 0) {
      // Prevent native scroll when pulling down at the top
      if (e.cancelable) e.preventDefault();
      
      const progress = Math.min(actualDiffY / PULL_THRESHOLD, 1.5); // Allow slight over-pull
      setPullProgress(Math.min(progress, 1));
      
      // Use requestAnimationFrame for smooth visual updates
      requestAnimationFrame(() => {
        if (contentRef.current) {
          // Apply a resistance curve
          const translateY = Math.min(actualDiffY * 0.4, PULL_THRESHOLD);
          contentRef.current.style.transform = `translateY(${translateY}px)`;
        }
      });
    }
  }, [isRefreshing]);

  const handleTouchEnd = useCallback(async () => {
    if (!isDragging.current) return;
    isDragging.current = false;
    
    if (isScrollingTrapped.current) {
      isScrollingTrapped.current = false;
      return;
    }

    const diff = currentY.current - startY.current;

    if (diff >= PULL_THRESHOLD && !isRefreshing && containerRef.current?.scrollTop === 0) {
      setIsRefreshing(true);
      setPullProgress(1);
      
      if (contentRef.current) {
        contentRef.current.style.transition = 'transform 0.3s cubic-bezier(0.2, 0.8, 0.2, 1)';
        contentRef.current.style.transform = `translateY(${PULL_THRESHOLD * 0.6}px)`;
      }

      try {
        await onRefresh();
      } finally {
        setIsRefreshing(false);
        setPullProgress(0);
        if (contentRef.current) {
          contentRef.current.style.transform = 'translateY(0px)';
          setTimeout(() => {
            if (contentRef.current) contentRef.current.style.transition = '';
          }, 300);
        }
      }
    } else {
      setPullProgress(0);
      if (contentRef.current) {
        contentRef.current.style.transition = 'transform 0.3s cubic-bezier(0.2, 0.8, 0.2, 1)';
        contentRef.current.style.transform = 'translateY(0px)';
        setTimeout(() => {
          if (contentRef.current) contentRef.current.style.transition = '';
        }, 300);
      }
    }
    
    startY.current = 0;
    currentY.current = 0;
  }, [isRefreshing, onRefresh]);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    container.addEventListener('touchstart', handleTouchStart, { passive: true });
    container.addEventListener('touchmove', handleTouchMove, { passive: false });
    container.addEventListener('touchend', handleTouchEnd, { passive: true });

    return () => {
      container.removeEventListener('touchstart', handleTouchStart);
      container.removeEventListener('touchmove', handleTouchMove);
      container.removeEventListener('touchend', handleTouchEnd);
    };
  }, [handleTouchStart, handleTouchMove, handleTouchEnd]);

  return (
    <div className="relative overflow-hidden h-full flex flex-col">
      <motion.div
        className="absolute top-0 left-0 right-0 flex justify-center items-center h-20 z-10 pointer-events-none"
        style={{
          opacity: isRefreshing ? 1 : pullProgress,
          y: isRefreshing ? 0 : -20 + (pullProgress * 20),
        }}
      >
        <div className="bg-white/80 backdrop-blur-md p-2 rounded-full shadow-lg border border-black/5">
          <motion.div
            animate={isRefreshing ? { rotate: 360 } : { rotate: pullProgress * 360 }}
            transition={isRefreshing ? { repeat: Infinity, duration: 1, ease: "linear" } : { type: "spring" }}
          >
            <RefreshCw className={`w-6 h-6 ${isRefreshing ? 'text-apple-blue' : 'text-apple-gray'}`} />
          </motion.div>
        </div>
      </motion.div>

      <div
        ref={containerRef}
        className="flex-1 overflow-y-auto overscroll-y-none custom-scrollbar"
        style={{ WebkitOverflowScrolling: 'touch' }}
      >
        <div ref={contentRef} className="min-h-full">
          {children}
        </div>
      </div>
    </div>
  );
};

export default PullToRefresh;
