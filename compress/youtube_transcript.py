#!/usr/bin/env python3
"""
YouTube Transcript Extractor using yt-dlp
Extracts transcripts from YouTube videos in multiple formats
yt-dlp is more reliable and robust against blocking
"""

import sys
import json
import argparse
import tempfile
import os
import re

# Import with error handling
try:
    import yt_dlp
except ImportError as e:
    print(json.dumps({"success": False, "error": f"Failed to import yt-dlp: {str(e)}. Please install with: pip install yt-dlp"}))
    sys.exit(1)

def get_ytdlp_options():
    """Get yt-dlp options with anti-bot measures based on official documentation
    
    According to yt-dlp docs:
    - Default clients: tv,android_sdkless,web (or android_sdkless,web_safari,web if no JS runtime)
    - android_sdkless is more reliable and less likely to be blocked
    - player_skip should be used carefully as it can cause missing formats/metadata
    """
    opts = {
        'quiet': True,
        'no_warnings': True,
        'skip_download': True,
        # Anti-bot measures
        'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'referer': 'https://www.youtube.com/',
        'extractor_args': {
            'youtube': {
                # Use android_sdkless first (more reliable, less bot detection)
                # Then fallback to web_safari and web
                # This matches the default behavior when no JS runtime is available
                'player_client': ['android_sdkless', 'web_safari', 'web'],
                # Don't skip webpage/configs aggressively - they may be needed for metadata
                # Only skip if we're still getting blocked (will try without skip first)
                # 'player_skip': [],  # Don't skip by default
            }
        },
        # Additional headers to appear more like a browser
        'http_headers': {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-us,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.7',
            'Keep-Alive': '300',
            'Connection': 'keep-alive',
        },
    }
    
    # Try to use cookies if available (from environment variable or file)
    cookies_path = os.environ.get('YOUTUBE_COOKIES_FILE')
    if cookies_path and os.path.exists(cookies_path):
        opts['cookiefile'] = cookies_path
    else:
        # Try common cookie file locations
        common_cookie_paths = [
            os.path.expanduser('~/.config/yt-dlp/cookies.txt'),
            os.path.expanduser('~/cookies.txt'),
            '/opt/backend/cookies.txt',
            os.path.join(os.path.dirname(__file__), 'cookies.txt'),
        ]
        for cookie_path in common_cookie_paths:
            if os.path.exists(cookie_path):
                opts['cookiefile'] = cookie_path
                break
    
    return opts

def format_time(seconds):
    """Convert seconds to SRT/VTT time format (HH:MM:SS,mmm)"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    millis = int((seconds % 1) * 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"

def format_time_vtt(seconds):
    """Convert seconds to VTT time format (HH:MM:SS.mmm)"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    millis = int((seconds % 1) * 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}.{millis:03d}"

def parse_vtt_time(time_str):
    """Parse VTT time string to seconds (HH:MM:SS.mmm)"""
    parts = time_str.split(':')
    if len(parts) == 3:
        hours = int(parts[0])
        minutes = int(parts[1])
        sec_parts = parts[2].split('.')
        seconds = int(sec_parts[0])
        milliseconds = int(sec_parts[1]) if len(sec_parts) > 1 else 0
        return hours * 3600 + minutes * 60 + seconds + milliseconds / 1000
    return 0

def parse_srt_time(time_str):
    """Parse SRT time string to seconds (HH:MM:SS,mmm)"""
    parts = time_str.split(':')
    if len(parts) == 3:
        hours = int(parts[0])
        minutes = int(parts[1])
        sec_parts = parts[2].split(',')
        seconds = int(sec_parts[0])
        milliseconds = int(sec_parts[1]) if len(sec_parts) > 1 else 0
        return hours * 3600 + minutes * 60 + seconds + milliseconds / 1000
    return 0

def parse_vtt_content(vtt_content):
    """Parse VTT subtitle content to list of entries"""
    entries = []
    lines = vtt_content.split('\n')
    i = 0
    
    # Skip WEBVTT header
    while i < len(lines) and not lines[i].strip().startswith('00:'):
        i += 1
    
    current_entry = None
    text_lines = []
    
    while i < len(lines):
        line = lines[i].strip()
        
        # Time range line (e.g., "00:00:00.000 --> 00:00:03.470")
        if '-->' in line:
            if current_entry and text_lines:
                current_entry['text'] = ' '.join(text_lines).strip()
                entries.append(current_entry)
            
            parts = line.split('-->')
            if len(parts) == 2:
                start_time = parse_vtt_time(parts[0].strip())
                end_time = parse_vtt_time(parts[1].strip())
                current_entry = {
                    'start': start_time,
                    'duration': end_time - start_time
                }
                text_lines = []
        elif line and current_entry:
            # Text line
            text_lines.append(line)
        
        i += 1
    
    # Add last entry
    if current_entry and text_lines:
        current_entry['text'] = ' '.join(text_lines).strip()
        entries.append(current_entry)
    
    return entries

def get_available_languages(video_id):
    """Get list of available transcript languages for a video using yt-dlp"""
    available_languages = []
    
    # Log to stderr for debugging
    import sys
    def log_debug(msg):
        print(f"DEBUG: {msg}", file=sys.stderr)
    
    try:
        video_url = f"https://www.youtube.com/watch?v={video_id}"
        
        ydl_opts = get_ytdlp_options()
        ydl_opts['listsubtitles'] = True
        ydl_opts['writesubtitles'] = False
        ydl_opts['writeautomaticsub'] = False
        
        log_debug(f"Getting available languages for video: {video_id}")
        if 'cookiefile' in ydl_opts:
            log_debug(f"Using cookies file: {ydl_opts['cookiefile']}")
        
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(video_url, download=False)
            
            # Check for subtitles
            if 'subtitles' in info:
                for lang_code, subtitles in info['subtitles'].items():
                    for sub in subtitles:
                        available_languages.append({
                            'code': lang_code,
                            'name': sub.get('name', lang_code),
                            'is_generated': False  # Manual subtitles
                        })
                        log_debug(f"Found manual subtitle: {lang_code} - {sub.get('name', lang_code)}")
            
            # Check for automatic captions
            if 'automatic_captions' in info:
                for lang_code, subtitles in info['automatic_captions'].items():
                    for sub in subtitles:
                        # Check if we already have this language
                        if not any(lang['code'] == lang_code for lang in available_languages):
                            available_languages.append({
                                'code': lang_code,
                                'name': sub.get('name', lang_code),
                                'is_generated': True  # Auto-generated
                            })
                            log_debug(f"Found auto subtitle: {lang_code} - {sub.get('name', lang_code)}")
        
        log_debug(f"Total available languages: {len(available_languages)}")
        
    except yt_dlp.utils.DownloadError as e:
        error_msg = str(e)
        log_debug(f"DownloadError getting available languages: {error_msg}")
        # Check if it's a bot detection error
        if 'bot' in error_msg.lower() or 'Sign in to confirm' in error_msg:
            log_debug("Bot detection detected when listing languages - may need cookies")
    except Exception as e:
        log_debug(f"Error getting available languages: {type(e).__name__}: {str(e)}")
        # Return empty list on error
    
    return available_languages

def get_transcript(video_id, language_codes=None, return_available=False):
    """Get transcript for a video using yt-dlp
    
    Args:
        video_id: YouTube video ID
        language_codes: List of language codes to try (e.g., ['es', 'en'])
        return_available: If True, also return available languages on error
    
    Returns:
        tuple: (transcript_data, available_languages) if return_available=True
        transcript_data if return_available=False
    """
    import sys
    def log_debug(msg):
        print(f"DEBUG: {msg}", file=sys.stderr)
    
    video_url = f"https://www.youtube.com/watch?v={video_id}"
    available_languages = []
    transcript_data = []
    
    log_debug(f"Getting transcript for video_id: {video_id}, languages: {language_codes}")
    
    try:
        # First, get available languages
        try:
            available_languages = get_available_languages(video_id)
        except Exception as e:
            log_debug(f"Error getting available languages: {e}")
        
        # Prepare language list to try
        langs_to_try = language_codes if language_codes else ['en']
        # Add English as fallback if not already in list
        if 'en' not in langs_to_try:
            langs_to_try.append('en')
        
        log_debug(f"Trying languages: {langs_to_try}")
        
        # Try each language
        for lang_code in langs_to_try:
            try:
                log_debug(f"Attempting to get transcript in language: {lang_code}")
                
                # Create temp directory for subtitle files
                with tempfile.TemporaryDirectory() as tmpdir:
                    ydl_opts = get_ytdlp_options()
                    ydl_opts['writesubtitles'] = True
                    ydl_opts['writeautomaticsub'] = True  # Also try auto-generated
                    ydl_opts['subtitleslangs'] = [lang_code]
                    ydl_opts['subtitlesformat'] = 'vtt'  # VTT is easier to parse
                    ydl_opts['outtmpl'] = os.path.join(tmpdir, '%(title)s.%(ext)s')
                    
                    if 'cookiefile' in ydl_opts:
                        log_debug(f"Using cookies file: {ydl_opts['cookiefile']}")
                    
                    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                        info = ydl.extract_info(video_url, download=True)
                        
                        # Find the downloaded subtitle file
                        subtitle_file = None
                        for file in os.listdir(tmpdir):
                            if file.endswith(f'.{lang_code}.vtt') or file.endswith(f'.{lang_code}.en.vtt'):
                                subtitle_file = os.path.join(tmpdir, file)
                                break
                        
                        # Also try without language code in filename
                        if not subtitle_file:
                            for file in os.listdir(tmpdir):
                                if file.endswith('.vtt'):
                                    subtitle_file = os.path.join(tmpdir, file)
                                    break
                        
                        if subtitle_file and os.path.exists(subtitle_file):
                            log_debug(f"Found subtitle file: {subtitle_file}")
                            with open(subtitle_file, 'r', encoding='utf-8') as f:
                                vtt_content = f.read()
                            
                            # Parse VTT content
                            transcript_data = parse_vtt_content(vtt_content)
                            log_debug(f"Parsed {len(transcript_data)} transcript entries")
                            
                            if transcript_data:
                                if return_available:
                                    return (transcript_data, available_languages)
                                return transcript_data
                        else:
                            log_debug(f"No subtitle file found for language: {lang_code}")
            
            except yt_dlp.utils.DownloadError as e:
                error_msg = str(e)
                log_debug(f"DownloadError fetching transcript in {lang_code}: {error_msg}")
                # Check if it's a bot detection error
                if 'bot' in error_msg.lower() or 'Sign in to confirm' in error_msg:
                    log_debug("Bot detection detected - trying fallback with player_skip...")
                    # Try once more with player_skip to reduce requests
                    try:
                        log_debug(f"Retrying {lang_code} with player_skip (reduced requests)...")
                        ydl_opts_fallback = get_ytdlp_options()
                        ydl_opts_fallback['extractor_args']['youtube']['player_skip'] = ['webpage', 'configs']
                        ydl_opts_fallback['writesubtitles'] = True
                        ydl_opts_fallback['writeautomaticsub'] = True
                        ydl_opts_fallback['subtitleslangs'] = [lang_code]
                        ydl_opts_fallback['subtitlesformat'] = 'vtt'
                        
                        with tempfile.TemporaryDirectory() as tmpdir_fallback:
                            ydl_opts_fallback['outtmpl'] = os.path.join(tmpdir_fallback, '%(title)s.%(ext)s')
                            with yt_dlp.YoutubeDL(ydl_opts_fallback) as ydl:
                                info = ydl.extract_info(video_url, download=True)
                                # Find subtitle file
                                subtitle_file = None
                                for file in os.listdir(tmpdir_fallback):
                                    if file.endswith('.vtt'):
                                        subtitle_file = os.path.join(tmpdir_fallback, file)
                                        break
                                if subtitle_file and os.path.exists(subtitle_file):
                                    log_debug(f"Found subtitle file with fallback: {subtitle_file}")
                                    with open(subtitle_file, 'r', encoding='utf-8') as f:
                                        vtt_content = f.read()
                                    transcript_data = parse_vtt_content(vtt_content)
                                    if transcript_data:
                                        if return_available:
                                            return (transcript_data, available_languages)
                                        return transcript_data
                    except Exception as fallback_error:
                        log_debug(f"Fallback attempt also failed: {fallback_error}")
                        log_debug("Bot detection persists - cookies are required")
                        log_debug("To fix: Export cookies from browser (Chrome/Firefox/etc) and set YOUTUBE_COOKIES_FILE environment variable")
                        log_debug("See: https://github.com/yt-dlp/yt-dlp/wiki/FAQ#how-do-i-pass-cookies-to-yt-dlp")
                continue
            except Exception as e:
                log_debug(f"Error fetching transcript in {lang_code}: {type(e).__name__}: {str(e)}")
                continue
        
        # If we get here, no transcript was found
        error_msg = "No transcript available for this video"
        if available_languages:
            lang_display = [f"{lang['name']} ({lang['code']})" for lang in available_languages[:5]]
            error_msg += f". However, {len(available_languages)} language(s) are available: {', '.join(lang_display)}"
        else:
            # Check if bot detection might be the issue
            error_msg += ". YouTube may be blocking requests. If you see 'Sign in to confirm you're not a bot' errors, please export cookies from your browser and set YOUTUBE_COOKIES_FILE environment variable. See: https://github.com/yt-dlp/yt-dlp/wiki/FAQ#how-do-i-pass-cookies-to-yt-dlp"
        
        if return_available:
            raise Exception(error_msg)
        raise Exception(error_msg)
    
    except Exception as e:
        log_debug(f"Unexpected error in get_transcript: {type(e).__name__}: {str(e)}")
        if return_available:
            raise Exception(f"Error fetching transcript: {str(e)}")
        raise Exception(f"Error fetching transcript: {str(e)}")

def format_as_text(transcript_data, include_timestamps=False):
    """Format transcript as plain text"""
    if include_timestamps:
        lines = []
        for item in transcript_data:
            start_time = format_time(item['start']).replace(',', ':')
            text = item['text'].strip()
            lines.append(f"[{start_time}] {text}")
        return '\n'.join(lines)
    else:
        return '\n'.join([item['text'].strip() for item in transcript_data])

def format_as_json(transcript_data):
    """Format transcript as JSON"""
    return json.dumps(transcript_data, indent=2, ensure_ascii=False)

def format_as_srt(transcript_data):
    """Format transcript as SRT subtitle format"""
    srt_lines = []
    for index, item in enumerate(transcript_data, start=1):
        start = item['start']
        duration = item.get('duration', 3.0)
        text = item['text'].strip().replace('\n', ' ')
        start_time = format_time(start)
        end_time = format_time(start + duration)
        srt_lines.append(f"{index}\n{start_time} --> {end_time}\n{text}\n")
    return '\n'.join(srt_lines)

def format_as_vtt(transcript_data):
    """Format transcript as VTT subtitle format"""
    vtt_lines = ["WEBVTT", ""]
    for item in transcript_data:
        start = item['start']
        duration = item.get('duration', 3.0)
        text = item['text'].strip().replace('\n', ' ')
        start_time = format_time_vtt(start)
        end_time = format_time_vtt(start + duration)
        vtt_lines.append(f"{start_time} --> {end_time}\n{text}")
    return '\n'.join(vtt_lines)

def main():
    parser = argparse.ArgumentParser(description='Extract YouTube video transcript using yt-dlp')
    parser.add_argument('video_id', help='YouTube video ID')
    parser.add_argument('--format', choices=['txt', 'txt-timestamps', 'json', 'srt', 'vtt'], 
                       default='txt', help='Output format')
    parser.add_argument('--language', help='Language code (e.g., en, es, fr)', default='en')
    
    args = parser.parse_args()
    
    # Log to stderr for debugging
    import sys
    def log_debug(msg):
        print(f"DEBUG: {msg}", file=sys.stderr)
    
    try:
        log_debug(f"=== Starting transcript extraction with yt-dlp ===")
        log_debug(f"Video ID: {args.video_id}")
        log_debug(f"Requested format: {args.format}")
        log_debug(f"Requested language: {args.language}")
        
        # Get transcript with available languages info
        language_list = [args.language] if args.language else ['en']
        if args.language != 'en':
            language_list.append('en')
        
        log_debug(f"Language list to try: {language_list}")
        
        available_languages = []
        try:
            transcript_data, available_languages = get_transcript(args.video_id, language_list, return_available=True)
            log_debug(f"Successfully got transcript with {len(transcript_data)} entries")
            log_debug(f"Available languages: {len(available_languages)}")
        except Exception as e:
            error_message = str(e)
            log_debug(f"get_transcript() raised exception: {type(e).__name__}: {error_message}")
            
            # Try to get available languages if not already retrieved
            if not available_languages:
                try:
                    available_languages = get_available_languages(args.video_id)
                except Exception:
                    available_languages = []
            
            error_result = {
                'success': False,
                'error': error_message,
                'video_id': args.video_id,
                'available_languages': available_languages
            }
            print(json.dumps(error_result))
            sys.exit(1)
        
        # Ensure transcript_data is a list
        if not isinstance(transcript_data, list):
            transcript_data = []
        
        if not transcript_data:
            error_result = {
                'success': False,
                'error': 'No transcript data retrieved',
                'video_id': args.video_id,
                'available_languages': available_languages
            }
            print(json.dumps(error_result))
            sys.exit(1)
        
        # Format based on requested format
        if args.format == 'txt':
            output = format_as_text(transcript_data, include_timestamps=False)
            content_type = 'text/plain'
        elif args.format == 'txt-timestamps':
            output = format_as_text(transcript_data, include_timestamps=True)
            content_type = 'text/plain'
        elif args.format == 'json':
            output = format_as_json(transcript_data)
            content_type = 'application/json'
        elif args.format == 'srt':
            output = format_as_srt(transcript_data)
            content_type = 'text/srt'
        elif args.format == 'vtt':
            output = format_as_vtt(transcript_data)
            content_type = 'text/vtt'
        
        # Get entries count
        entries_count = len(transcript_data)
        
        # Return as JSON for API
        result = {
            'success': True,
            'format': args.format,
            'content': output,
            'content_type': content_type,
            'video_id': args.video_id,
            'entries_count': entries_count,
            'available_languages': available_languages
        }
        
        print(json.dumps(result))
        
    except Exception as e:
        try:
            available_languages = get_available_languages(args.video_id)
        except:
            available_languages = []
        error_result = {
            'success': False,
            'error': str(e),
            'video_id': args.video_id,
            'available_languages': available_languages
        }
        print(json.dumps(error_result))
        sys.exit(1)

if __name__ == '__main__':
    main()
