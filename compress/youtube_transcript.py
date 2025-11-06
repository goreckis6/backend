#!/usr/bin/env python3
"""
YouTube Transcript Extractor
Extracts transcripts from YouTube videos in multiple formats
"""

import sys
import json
import argparse

# Import with error handling
try:
    from youtube_transcript_api import YouTubeTranscriptApi
    from youtube_transcript_api._errors import (
        TranscriptsDisabled, 
        NoTranscriptFound, 
        VideoUnavailable,
        TooManyRequests,
        YouTubeRequestFailed,
        CouldNotRetrieveTranscript
    )
except ImportError as e:
    print(json.dumps({"success": False, "error": f"Failed to import youtube_transcript_api: {str(e)}"}))
    sys.exit(1)

# Try to import RequestBlocked if available (may not exist in all versions)
try:
    from youtube_transcript_api._errors import RequestBlocked, IpBlocked
    HAS_BLOCKING_ERRORS = True
except ImportError:
    HAS_BLOCKING_ERRORS = False
    RequestBlocked = type('RequestBlocked', (Exception,), {})
    IpBlocked = type('IpBlocked', (Exception,), {})

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

def get_available_languages(video_id):
    """Get list of available transcript languages for a video"""
    ytt_api = YouTubeTranscriptApi()
    available_languages = []
    
    try:
        transcript_list = ytt_api.list(video_id)
        # transcript_list might be iterable, try to iterate over it
        try:
            # Try to iterate
            for transcript in transcript_list:
                try:
                    # Try to access transcript attributes
                    lang_code = getattr(transcript, 'language_code', None) or getattr(transcript, 'code', None)
                    lang_name = getattr(transcript, 'language', None) or getattr(transcript, 'name', None)
                    is_gen = getattr(transcript, 'is_generated', False)
                    
                    if lang_code and lang_name:
                        available_languages.append({
                            'code': lang_code,
                            'name': lang_name,
                            'is_generated': is_gen
                        })
                except Exception:
                    # Skip individual transcript errors, continue with others
                    continue
        except TypeError:
            # transcript_list might not be iterable directly, try find methods
            try:
                # Try common languages
                for lang_code in ['en', 'es', 'fr', 'de', 'it', 'pt', 'ru', 'ja', 'ko', 'zh', 'sv', 'no', 'da', 'fi', 'nl']:
                    try:
                        transcript = transcript_list.find_transcript([lang_code])
                        available_languages.append({
                            'code': lang_code,
                            'name': transcript.language if hasattr(transcript, 'language') else lang_code,
                            'is_generated': getattr(transcript, 'is_generated', False)
                        })
                    except Exception:
                        continue
            except Exception:
                pass
    except Exception as e:
        # If list() fails, return empty list - we'll try to populate it during fetch
        # Could be VideoUnavailable, TranscriptsDisabled, or network issues
        pass
    
    return available_languages

def get_transcript(video_id, language_codes=None, return_available=False):
    """Get transcript for a video, trying multiple languages
    
    Args:
        video_id: YouTube video ID
        language_codes: List of language codes to try (e.g., ['es', 'en'])
        return_available: If True, also return available languages on error
    
    Returns:
        tuple: (transcript_data, available_languages) if return_available=True
        transcript_data if return_available=False
    """
    # According to documentation, we need to create an instance first
    ytt_api = YouTubeTranscriptApi()
    available_languages = []
    transcript_list_obj = None
    last_error = None
    
    # Log to stderr for debugging (will be visible in backend logs)
    import sys
    def log_debug(msg):
        print(f"DEBUG: {msg}", file=sys.stderr)
    
    log_debug(f"Getting transcript for video_id: {video_id}, languages: {language_codes}")
    
    try:
        # First, try to get available languages by listing transcripts
        # This also gives us the transcript_list object for later use
        try:
            log_debug("Attempting to list transcripts...")
            transcript_list_obj = ytt_api.list(video_id)
            log_debug(f"Successfully got transcript_list_obj: {type(transcript_list_obj)}")
            for transcript in transcript_list_obj:
                try:
                    lang_code = transcript.language_code
                    lang_name = transcript.language
                    is_gen = transcript.is_generated
                    available_languages.append({
                        'code': lang_code,
                        'name': lang_name,
                        'is_generated': is_gen
                    })
                    log_debug(f"Found language: {lang_name} ({lang_code}), generated: {is_gen}")
                except Exception as e:
                    last_error = f"Error processing transcript in list: {str(e)}"
                    log_debug(f"Error processing transcript: {e}")
                    continue
            log_debug(f"Total available languages found: {len(available_languages)}")
        except Exception as e:
            # If list() fails, we'll try other methods and populate available_languages later
            last_error = f"list() failed: {type(e).__name__}: {str(e)}"
            log_debug(f"list() failed: {last_error}")
            # Try to continue anyway - sometimes list() fails but fetch() works
            pass
        
        # Try with preferred languages first using fetch()
        if language_codes:
            try:
                log_debug(f"Trying fetch() with languages: {language_codes}")
                fetched_transcript = ytt_api.fetch(video_id, languages=language_codes)
                log_debug(f"Successfully fetched transcript with fetch()")
                data = fetched_transcript.to_raw_data()
                log_debug(f"Got {len(data)} transcript entries")
                # If we got available_languages from list(), use them, otherwise try to get them now
                if not available_languages and transcript_list_obj is None:
                    try:
                        transcript_list_obj = ytt_api.list(video_id)
                        for transcript in transcript_list_obj:
                            try:
                                available_languages.append({
                                    'code': transcript.language_code,
                                    'name': transcript.language,
                                    'is_generated': transcript.is_generated
                                })
                            except Exception:
                                continue
                    except Exception:
                        pass
                if return_available:
                    return (data, available_languages)
                return data
            except (NoTranscriptFound, TranscriptsDisabled) as e:
                # These are expected exceptions, continue to next method
                last_error = f"{type(e).__name__}: {str(e)}"
                log_debug(f"fetch() with {language_codes} failed: {last_error}")
                pass
            except Exception as e:
                # Other exceptions, continue to next method
                last_error = f"{type(e).__name__}: {str(e)}"
                log_debug(f"fetch() with {language_codes} raised exception: {last_error}")
                pass
        
        # Try English as fallback using fetch()
        try:
            log_debug("Trying fetch() with English as fallback...")
            fetched_transcript = ytt_api.fetch(video_id, languages=['en'])
            log_debug("Successfully fetched English transcript")
            data = fetched_transcript.to_raw_data()
            log_debug(f"Got {len(data)} transcript entries")
            # Populate available_languages if we haven't yet
            if not available_languages and transcript_list_obj is None:
                try:
                    transcript_list_obj = ytt_api.list(video_id)
                    for transcript in transcript_list_obj:
                        try:
                            available_languages.append({
                                'code': transcript.language_code,
                                'name': transcript.language,
                                'is_generated': transcript.is_generated
                            })
                        except Exception:
                            continue
                except Exception:
                    pass
            if return_available:
                return (data, available_languages)
            return data
        except (NoTranscriptFound, TranscriptsDisabled) as e:
            last_error = f"{type(e).__name__}: {str(e)}"
            log_debug(f"fetch() with English failed: {last_error}")
            pass
        except Exception as e:
            last_error = f"{type(e).__name__}: {str(e)}"
            log_debug(f"fetch() with English raised exception: {last_error}")
            pass
        
        # Try without language specification (defaults to English) using fetch()
        try:
            log_debug("Trying fetch() without language specification...")
            fetched_transcript = ytt_api.fetch(video_id)
            log_debug("Successfully fetched transcript without language")
            data = fetched_transcript.to_raw_data()
            log_debug(f"Got {len(data)} transcript entries")
            # Populate available_languages if we haven't yet
            if not available_languages and transcript_list_obj is None:
                try:
                    transcript_list_obj = ytt_api.list(video_id)
                    for transcript in transcript_list_obj:
                        try:
                            available_languages.append({
                                'code': transcript.language_code,
                                'name': transcript.language,
                                'is_generated': transcript.is_generated
                            })
                        except Exception:
                            continue
                except Exception:
                    pass
            if return_available:
                return (data, available_languages)
            return data
        except (NoTranscriptFound, TranscriptsDisabled) as e:
            last_error = f"{type(e).__name__}: {str(e)}"
            log_debug(f"fetch() without language failed: {last_error}")
            pass
        except Exception as e:
            last_error = f"{type(e).__name__}: {str(e)}"
            log_debug(f"fetch() without language raised exception: {last_error}")
            pass
        
        # If direct fetch fails, try using list() to find available transcripts
        # If we already have transcript_list_obj, use it; otherwise get it
        if transcript_list_obj is None:
            try:
                transcript_list_obj = ytt_api.list(video_id)
                # Populate available_languages from transcript_list_obj
                if not available_languages:
                    for transcript in transcript_list_obj:
                        try:
                            available_languages.append({
                                'code': transcript.language_code,
                                'name': transcript.language,
                                'is_generated': transcript.is_generated
                            })
                        except Exception:
                            continue
            except Exception as e:
                # If list() fails completely, we can't get available languages
                pass
        
        if transcript_list_obj:
            # Try to find transcript in preferred languages using list()
            if language_codes:
                try:
                    transcript = transcript_list_obj.find_transcript(language_codes)
                    fetched_transcript = transcript.fetch()
                    data = fetched_transcript.to_raw_data()
                    if return_available:
                        return (data, available_languages)
                    return data
                except (NoTranscriptFound, TranscriptsDisabled) as e:
                    last_error = f"{type(e).__name__}: {str(e)}"
                    pass
                except Exception as e:
                    last_error = f"{type(e).__name__}: {str(e)}"
                    pass
            
            # Try English using list()
            try:
                transcript = transcript_list_obj.find_transcript(['en'])
                fetched_transcript = transcript.fetch()
                data = fetched_transcript.to_raw_data()
                if return_available:
                    return (data, available_languages)
                return data
            except (NoTranscriptFound, TranscriptsDisabled) as e:
                last_error = f"{type(e).__name__}: {str(e)}"
                pass
            except Exception as e:
                last_error = f"{type(e).__name__}: {str(e)}"
                pass
            
            # Get any available transcript from list()
            try:
                for transcript in transcript_list_obj:
                    try:
                        fetched_transcript = transcript.fetch()
                        data = fetched_transcript.to_raw_data()
                        if return_available:
                            return (data, available_languages)
                        return data
                    except Exception as e:
                        # Continue to next transcript
                        last_error = f"{type(e).__name__}: {str(e)}"
                        continue
            except Exception as e:
                last_error = f"{type(e).__name__}: {str(e)}"
                pass
        
        # If we get here, no transcript was found
        # But we should have available_languages populated by now
        log_debug(f"All fetch() methods failed. Last error: {last_error}")
        log_debug(f"Available languages found: {len(available_languages)}")
        if available_languages:
            lang_list = [f"{lang['name']} ({lang['code']})" for lang in available_languages]
            log_debug(f"Available languages: {lang_list}")
        
        error_msg = "No transcript available for this video"
        if last_error:
            error_msg += f" (Last error: {last_error})"
        if available_languages:
            lang_display = [f"{lang['name']} ({lang['code']})" for lang in available_languages[:5]]
            error_msg += f". However, {len(available_languages)} language(s) are available: {', '.join(lang_display)}"
        if return_available:
            raise Exception(error_msg)
        raise Exception(error_msg)
        
    except TranscriptsDisabled as e:
        log_debug(f"TranscriptsDisabled exception: {str(e)}")
        if return_available:
            raise Exception("Transcripts are disabled for this video")
        raise Exception("Transcripts are disabled for this video")
    except NoTranscriptFound as e:
        log_debug(f"NoTranscriptFound exception: {str(e)}")
        if return_available:
            raise Exception("No transcript found for this video")
        raise Exception("No transcript found for this video")
    except VideoUnavailable as e:
        log_debug(f"VideoUnavailable exception: {str(e)}")
        if return_available:
            raise Exception("Video is unavailable or doesn't exist")
        raise Exception("Video is unavailable or doesn't exist")
    except Exception as e:
        # Check if this is an IP blocking error (if available in this version)
        if HAS_BLOCKING_ERRORS and isinstance(e, (RequestBlocked, IpBlocked)):
            log_debug(f"IP blocking detected: {type(e).__name__}: {str(e)}")
            error_msg = "YouTube is blocking requests from this server. This is a known issue with cloud providers. Please try again later or contact support."
            if return_available:
                raise Exception(error_msg)
            raise Exception(error_msg)
        
        # Check for other known YouTube API errors
        if isinstance(e, (TooManyRequests, YouTubeRequestFailed, CouldNotRetrieveTranscript)):
            log_debug(f"YouTube API error: {type(e).__name__}: {str(e)}")
            error_msg = f"YouTube API error: {str(e)}"
            if return_available:
                raise Exception(error_msg)
            raise Exception(error_msg)
        
        # Check if error message suggests IP blocking
        error_msg = str(e)
        if 'blocked' in error_msg.lower() or 'request blocked' in error_msg.lower():
            log_debug(f"IP blocking detected (from error message): {error_msg}")
            error_msg = "YouTube is blocking requests from this server. Please try again later."
        
        log_debug(f"Unexpected exception in get_transcript: {type(e).__name__}: {error_msg}")
        if return_available:
            raise Exception(f"Error fetching transcript: {error_msg}")
        raise Exception(f"Error fetching transcript: {error_msg}")

def format_as_text(transcript_data, include_timestamps=False):
    """Format transcript as plain text"""
    # transcript_data should be a list of dicts from to_raw_data()
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
    # transcript_data should already be a list of dicts from to_raw_data()
    return json.dumps(transcript_data, indent=2, ensure_ascii=False)

def format_as_srt(transcript_data):
    """Format transcript as SRT subtitle format"""
    # transcript_data should already be a list of dicts from to_raw_data()
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
    # transcript_data should already be a list of dicts from to_raw_data()
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
    parser = argparse.ArgumentParser(description='Extract YouTube video transcript')
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
        log_debug(f"=== Starting transcript extraction ===")
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
            
            # If available_languages is empty, try multiple methods to get them
            if not available_languages:
                # Method 1: Try get_available_languages function
                try:
                    available_languages = get_available_languages(args.video_id)
                except Exception as e2:
                    # If that fails, try direct API call
                    try:
                        ytt_api = YouTubeTranscriptApi()
                        transcript_list = ytt_api.list(args.video_id)
                        for transcript in transcript_list:
                            try:
                                lang_code = getattr(transcript, 'language_code', None) or getattr(transcript, 'code', None)
                                lang_name = getattr(transcript, 'language', None) or getattr(transcript, 'name', None)
                                is_gen = getattr(transcript, 'is_generated', False)
                                if lang_code:
                                    available_languages.append({
                                        'code': lang_code,
                                        'name': lang_name or lang_code,
                                        'is_generated': is_gen
                                    })
                            except Exception:
                                continue
                    except Exception as e3:
                        # Last attempt: try to fetch without language to see if any transcript exists
                        try:
                            # Try fetching with no language to get default
                            fetched = ytt_api.fetch(args.video_id)
                            # If we get here, transcript exists but we couldn't list languages
                            # At least we know transcripts are available
                            error_message = f"Transcript exists but couldn't list languages. Original error: {error_message}"
                        except Exception as e4:
                            # All methods failed
                            pass
            
            error_result = {
                'success': False,
                'error': error_message,
                'video_id': args.video_id,
                'available_languages': available_languages
            }
            print(json.dumps(error_result))
            sys.exit(1)
        
        # get_transcript now returns a list of dicts (from to_raw_data())
        # So transcript_data should already be in the correct format
        if not isinstance(transcript_data, list):
            # Fallback: if somehow it's not a list, try to convert
            try:
                if hasattr(transcript_data, '__iter__'):
                    transcript_data = list(transcript_data)
                else:
                    transcript_data = [transcript_data] if transcript_data else []
            except Exception:
                transcript_data = []
        
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
        
        # Get entries count - transcript_data should be a list
        entries_count = len(transcript_data) if isinstance(transcript_data, list) else 0
        
        # Get available languages for this video
        available_languages = get_available_languages(args.video_id)
        
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
        
    except TranscriptsDisabled:
        try:
            available_languages = get_available_languages(args.video_id)
        except:
            available_languages = []
        error_result = {
            'success': False,
            'error': 'Transcripts are disabled for this video',
            'video_id': args.video_id,
            'available_languages': available_languages
        }
        print(json.dumps(error_result))
        sys.exit(1)
    except NoTranscriptFound:
        try:
            available_languages = get_available_languages(args.video_id)
        except:
            available_languages = []
        error_result = {
            'success': False,
            'error': 'No transcript found for this video in the specified language(s)',
            'video_id': args.video_id,
            'available_languages': available_languages
        }
        print(json.dumps(error_result))
        sys.exit(1)
    except VideoUnavailable:
        error_result = {
            'success': False,
            'error': 'Video is unavailable or doesn\'t exist',
            'video_id': args.video_id,
            'available_languages': []
        }
        print(json.dumps(error_result))
        sys.exit(1)
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
