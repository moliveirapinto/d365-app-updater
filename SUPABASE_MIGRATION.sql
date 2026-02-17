-- ═══════════════════════════════════════════════════════════
-- Optional Migration: Add display columns to update_schedules
-- ═══════════════════════════════════════════════════════════
--
-- This migration adds optional columns to store the user's
-- original timezone selection. This makes the UI more accurate
-- when daylight saving time changes occur.
--
-- HOW TO RUN:
-- 1. Go to your Supabase project dashboard
-- 2. Navigate to the SQL Editor
-- 3. Paste and run this script
--
-- Note: The app works fine WITHOUT these columns (it calculates
-- them on the fly), but having them prevents minor DST-related
-- display inconsistencies.
-- ═══════════════════════════════════════════════════════════

-- Add display columns (if they don't already exist)
DO $$ 
BEGIN
    -- Add display_day column
    IF NOT EXISTS (
        SELECT FROM information_schema.columns 
        WHERE table_name = 'update_schedules' 
        AND column_name = 'display_day'
    ) THEN
        ALTER TABLE update_schedules 
        ADD COLUMN display_day integer;
        
        COMMENT ON COLUMN update_schedules.display_day IS 
        'User''s original day selection (0=Sunday, 6=Saturday) in their local timezone';
    END IF;
    
    -- Add display_time column
    IF NOT EXISTS (
        SELECT FROM information_schema.columns 
        WHERE table_name = 'update_schedules' 
        AND column_name = 'display_time'
    ) THEN
        ALTER TABLE update_schedules 
        ADD COLUMN display_time text;
        
        COMMENT ON COLUMN update_schedules.display_time IS 
        'User''s original time selection (HH:MM format) in their local timezone';
    END IF;
END $$;

-- Verify the changes
SELECT column_name, data_type, column_default, is_nullable
FROM information_schema.columns
WHERE table_name = 'update_schedules'
ORDER BY ordinal_position;

-- Migration completed successfully!
-- The app will now store both UTC values (for execution) and
-- display values (for showing the user what they selected).
