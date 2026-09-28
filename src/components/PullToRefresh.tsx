import React, { useState, useRef, useEffect, useCallback } from 'react';
import { motion } from 'motion/react';
import { RefreshCw } from 'lucide-react';

interface PullToRefreshProps {
  onRefresh: () => Promise<void>;
  children: React.ReactNode;
  disabled?: boolean;
}

const PullToRefresh: React.FC<PullToRefreshProps> = ({ onRefresh, children, disabled = false }) => {
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [pullProgress, setPullProgress] = useState(0);
  const containerRef = useRef<HTMLDivElement>(null);
  const contentRef = useRef<HTMLDivElement>(null);

  const startX = useRef(0);
  const startY = useRef(0);
  const currentY = useRef(0);
  const isDragging = useRef(false);
  const isAborted = useRef(false);
  const PULL_THRESHOLD = 80;

  // Check if viewport is truly at the top of the page
  const isAtTop = () => {
    const winY = window.scrollY || window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0;
    const containerY = containerRef.current?.scrollTop || 0;
    return winY <= 1 && containerY <= 1;
  };

  const handleTouchStart = useCallback((e: TouchEvent) => {
    if (disabled || isRefreshing) return;
    
    // ONLY initiate drag if user touches when already at the absolute top of the page
    if (isAtTop()) {
      startX.current = e.touches[0].clientX;
      startY.current = e.touches[0].clientY;
      currentY.current = e.touches[0].clientY;
      isDragging.current = true;
      isAborted.current = false;
    } else {
      isDragging.current = false;
      isAborted.current = true;
    }
  }, [disabled, isRefreshing]);

  const handleTouchMove = useCallback((e: TouchEvent) => {
    if (!isDragging.current || isAborted.current || isRefreshing || disabled) return;

    const currentX = e.touches[0].clientX;
    const currentYVal = e.touches[0].clientY;
    currentY.current = currentYVal;
    
    const diffX = Math.abs(currentX - startX.current);
    const actualDiffY = currentYVal - startY.current;

    // If swiping upwards (scrolling down into page), ABORT IMMEDIATELY!
    // Never trap or interfere with downward page scrolling
    if (actualDiffY < 0) {
      isDragging.current = false;
      isAborted.current = true;
      setPullProgress(0);
      if (contentRef.current) contentRef.current.style.transform = '';
      return;
    }

    // If horizontal motion exceeds vertical motion, user is swiping horizontally (tabs, carousel, swipeable item)
    if (diffX > actualDiffY && diffX > 8) {
      isDragging.current = false;
      isAborted.current = true;
      setPullProgress(0);
      if (contentRef.current) contentRef.current.style.transform = '';
      return;
    }

    // Double check that we didn't scroll down
    if (!isAtTop()) {
      isDragging.current = false;
      isAborted.current = true;
      setPullProgress(0);
      if (contentRef.current) contentRef.current.style.transform = '';
      return;
    }

    // Only when pulling down from the top:
    if (actualDiffY > 0) {
      // Prevent browser default overscroll only when we are genuinely handling the pull
      if (e.cancelable) e.preventDefault();
      
      const progress = Math.min(actualDiffY / PULL_THRESHOLD, 1.5);
      setPullProgress(Math.min(progress, 1));
      
      if (contentRef.current) {
        const translateY = Math.min(actualDiffY * 0.4, PULL_THRESHOLD);
        contentRef.current.style.transform = `translateY(${translateY}px)`;
      }
    }
  }, [disabled, isRefreshing]);

  const handleTouchEnd = useCallback(async () => {
    if (!isDragging.current || isAborted.current || disabled) {
      isDragging.current = false;
      isAborted.current = false;
      return;
    }
    isDragging.current = false;

    const diff = currentY.current - startY.current;

    if (diff >= PULL_THRESHOLD && !isRefreshing && isAtTop()) {
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
            if (contentRef.current) {
              contentRef.current.style.transition = '';
              contentRef.current.style.transform = '';
            }
          }, 300);
        }
      }
    } else {
      setPullProgress(0);
      if (contentRef.current) {
        contentRef.current.style.transition = 'transform 0.3s cubic-bezier(0.2, 0.8, 0.2, 1)';
        contentRef.current.style.transform = 'translateY(0px)';
        setTimeout(() => {
          if (contentRef.current) {
            contentRef.current.style.transition = '';
            contentRef.current.style.transform = '';
          }
        }, 300);
      }
    }
    
    startY.current = 0;
    currentY.current = 0;
  }, [disabled, isRefreshing, onRefresh]);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    container.addEventListener('touchstart', handleTouchStart, { passive: true });
    container.addEventListener('touchmove', handleTouchMove, { passive: false });
    container.addEventListener('touchend', handleTouchEnd, { passive: true });
    container.addEventListener('touchcancel', handleTouchEnd, { passive: true });

    return () => {
      container.removeEventListener('touchstart', handleTouchStart);
      container.removeEventListener('touchmove', handleTouchMove);
      container.removeEventListener('touchend', handleTouchEnd);
      container.removeEventListener('touchcancel', handleTouchEnd);
    };
  }, [handleTouchStart, handleTouchMove, handleTouchEnd]);

  return (
    <div className="relative w-full flex-1 flex flex-col">
      <motion.div
        className="absolute top-0 left-0 right-0 flex justify-center items-center h-16 z-20 pointer-events-none"
        style={{
          opacity: isRefreshing ? 1 : pullProgress,
          y: isRefreshing ? 0 : -20 + (pullProgress * 20),
        }}
      >
        <div className="bg-white/90 backdrop-blur-md p-2.5 rounded-full shadow-lg border border-black/5">
          <motion.div
            animate={isRefreshing ? { rotate: 360 } : { rotate: pullProgress * 360 }}
            transition={isRefreshing ? { repeat: Infinity, duration: 1, ease: "linear" } : { type: "spring" }}
          >
            <RefreshCw className={`w-5 h-5 ${isRefreshing ? 'text-apple-blue' : 'text-apple-gray'}`} />
          </motion.div>
        </div>
      </motion.div>

      <div
        ref={containerRef}
        className="w-full flex-1"
      >
        <div ref={contentRef} className="w-full min-h-full">
          {children}
        </div>
      </div>
    </div>
  );
};

export default PullToRefresh;
