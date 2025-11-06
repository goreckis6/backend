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
    from youtube_transcript_api._errors import TranscriptsDisabled, NoTranscriptFound, VideoUnavailable
    # Also try importing the module to check for module-level functions
    import youtube_transcript_api as yt_api_module
except ImportError as e:
    print(json.dumps({"success": False, "error": f"Failed to import youtube_transcript_api: {str(e)}"}))
    sys.exit(1)

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

def get_transcript(video_id, language_codes=None):
    """Get transcript for a video, trying multiple languages"""
    # In version 1.2.3, the API uses 'list' and 'fetch' methods
    # Try different patterns to find the correct API usage
    
    # Pattern 1: Try fetch() as a class method
    try:
        if language_codes:
            for lang_code in language_codes:
                try:
                    transcript = YouTubeTranscriptApi.fetch(video_id, languages=[lang_code])
                    return transcript
                except (NoTranscriptFound, TranscriptsDisabled):
                    continue
                except Exception:
                    continue
        
        try:
            transcript = YouTubeTranscriptApi.fetch(video_id, languages=['en'])
            return transcript
        except (NoTranscriptFound, TranscriptsDisabled):
            pass
        except Exception:
            pass
        
        try:
            transcript = YouTubeTranscriptApi.fetch(video_id)
            return transcript
        except (NoTranscriptFound, TranscriptsDisabled):
            pass
        except Exception:
            pass
    except Exception:
        pass
    
    # Pattern 2: Try creating an instance and calling methods on it
    try:
        api_instance = YouTubeTranscriptApi()
        
        # Try fetch on instance
        if language_codes:
            for lang_code in language_codes:
                try:
                    transcript = api_instance.fetch(video_id, languages=[lang_code])
                    return transcript
                except (NoTranscriptFound, TranscriptsDisabled):
                    continue
                except Exception:
                    continue
        
        try:
            transcript = api_instance.fetch(video_id, languages=['en'])
            return transcript
        except (NoTranscriptFound, TranscriptsDisabled):
            pass
        except Exception:
            pass
        
        try:
            transcript = api_instance.fetch(video_id)
            return transcript
        except (NoTranscriptFound, TranscriptsDisabled):
            pass
        except Exception:
            pass
        
        # Try list on instance (this should work if list is an instance method)
        try:
            # Check if list is callable and what its signature is
            if hasattr(api_instance, 'list') and callable(getattr(api_instance, 'list')):
                import inspect
                try:
                    sig = inspect.signature(api_instance.list)
                    # If it requires video_id, call it with video_id
                    transcript_list = api_instance.list(video_id)
                except TypeError:
                    # Maybe list doesn't take video_id, try without
                    try:
                        transcript_list = api_instance.list()
                    except:
                        # Maybe list is a property, not a method
                        transcript_list = api_instance.list
                except Exception:
                    transcript_list = api_instance.list(video_id)
            else:
                transcript_list = api_instance.list(video_id)
            
            # Try preferred languages
            if language_codes:
                for lang_code in language_codes:
                    try:
                        if hasattr(transcript_list, '__iter__'):
                            for item in transcript_list:
                                try:
                                    if hasattr(item, 'language_code') and item.language_code == lang_code:
                                        return item.fetch()
                                    if hasattr(item, 'language') and item.language == lang_code:
                                        return item.fetch()
                                except:
                                    continue
                    except:
                        continue
            
            # Try English
            try:
                if hasattr(transcript_list, '__iter__'):
                    for item in transcript_list:
                        try:
                            if hasattr(item, 'language_code') and item.language_code == 'en':
                                return item.fetch()
                            if hasattr(item, 'language') and item.language == 'en':
                                return item.fetch()
                        except:
                            continue
            except:
                pass
            
            # Get any available
            try:
                if hasattr(transcript_list, '__iter__'):
                    for item in transcript_list:
                        try:
                            return item.fetch()
                        except:
                            continue
            except:
                pass
            
            if isinstance(transcript_list, list) and len(transcript_list) > 0:
                return transcript_list[0].fetch()
        except Exception:
            pass
    except Exception:
        pass
    
    # Pattern 3: Try module-level functions
    try:
        # Check if fetch and list are module-level functions
        if hasattr(yt_api_module, 'fetch'):
            if language_codes:
                for lang_code in language_codes:
                    try:
                        transcript = yt_api_module.fetch(video_id, languages=[lang_code])
                        return transcript
                    except (NoTranscriptFound, TranscriptsDisabled):
                        continue
                    except Exception:
                        continue
            
            try:
                transcript = yt_api_module.fetch(video_id, languages=['en'])
                return transcript
            except (NoTranscriptFound, TranscriptsDisabled):
                pass
            except Exception:
                pass
            
            try:
                transcript = yt_api_module.fetch(video_id)
                return transcript
            except (NoTranscriptFound, TranscriptsDisabled):
                pass
            except Exception:
                pass
        
        if hasattr(yt_api_module, 'list'):
            transcript_list = yt_api_module.list(video_id)
            
            if language_codes:
                for lang_code in language_codes:
                    try:
                        if hasattr(transcript_list, '__iter__'):
                            for item in transcript_list:
                                try:
                                    if hasattr(item, 'language_code') and item.language_code == lang_code:
                                        return item.fetch()
                                    if hasattr(item, 'language') and item.language == lang_code:
                                        return item.fetch()
                                except:
                                    continue
                    except:
                        continue
            
            try:
                if hasattr(transcript_list, '__iter__'):
                    for item in transcript_list:
                        try:
                            if hasattr(item, 'language_code') and item.language_code == 'en':
                                return item.fetch()
                            if hasattr(item, 'language') and item.language == 'en':
                                return item.fetch()
                        except:
                            continue
            except:
                pass
            
            try:
                if hasattr(transcript_list, '__iter__'):
                    for item in transcript_list:
                        try:
                            return item.fetch()
                        except:
                            continue
            except:
                pass
            
            if isinstance(transcript_list, list) and len(transcript_list) > 0:
                return transcript_list[0].fetch()
    except Exception:
        pass
    
    # Pattern 4: Maybe the API works differently - try the old get_transcript pattern
    # Some versions might still support it
    try:
        # Check if there's a get_transcript function somewhere
        if hasattr(yt_api_module, 'get_transcript'):
            if language_codes:
                for lang_code in language_codes:
                    try:
                        transcript = yt_api_module.get_transcript(video_id, languages=[lang_code])
                        return transcript
                    except (NoTranscriptFound, TranscriptsDisabled):
                        continue
                    except Exception:
                        continue
            
            try:
                return yt_api_module.get_transcript(video_id, languages=['en'])
            except (NoTranscriptFound, TranscriptsDisabled):
                pass
            except Exception:
                pass
            
            try:
                return yt_api_module.get_transcript(video_id)
            except (NoTranscriptFound, TranscriptsDisabled):
                pass
            except Exception:
                pass
    except Exception:
        pass
    
    # If nothing works, raise error with helpful message
    raise Exception("Unable to fetch transcript. The API methods 'fetch' and 'list' are available but couldn't be called successfully. Please check the youtube-transcript-api documentation for version 1.2.3.")

def format_as_text(transcript_data, include_timestamps=False):
    """Format transcript as plain text"""
    # Handle both list of dicts and FetchedTranscriptSnippet object
    if not isinstance(transcript_data, list):
        # If it's an object, try to convert it to a list
        if hasattr(transcript_data, '__iter__'):
            transcript_data = list(transcript_data)
        elif hasattr(transcript_data, 'text') and hasattr(transcript_data, 'start'):
            # Single item object
            transcript_data = [transcript_data]
        else:
            # Try to get data attribute
            if hasattr(transcript_data, 'data'):
                transcript_data = transcript_data.data
            elif hasattr(transcript_data, 'transcript'):
                transcript_data = transcript_data.transcript
    
    if include_timestamps:
        lines = []
        for item in transcript_data:
            # Handle both dict and object
            if isinstance(item, dict):
                start_time = format_time(item['start']).replace(',', ':')
                text = item['text'].strip()
            else:
                # Object with attributes
                start = getattr(item, 'start', getattr(item, 'timestamp', 0))
                text = getattr(item, 'text', getattr(item, 'content', '')).strip()
                start_time = format_time(start).replace(',', ':')
            lines.append(f"[{start_time}] {text}")
        return '\n'.join(lines)
    else:
        result = []
        for item in transcript_data:
            if isinstance(item, dict):
                result.append(item['text'].strip())
            else:
                text = getattr(item, 'text', getattr(item, 'content', '')).strip()
                result.append(text)
        return '\n'.join(result)

def format_as_json(transcript_data):
    """Format transcript as JSON"""
    # Convert object to list of dicts if needed
    if not isinstance(transcript_data, list):
        if hasattr(transcript_data, '__iter__'):
            transcript_data = list(transcript_data)
        elif hasattr(transcript_data, 'text') and hasattr(transcript_data, 'start'):
            transcript_data = [transcript_data]
        elif hasattr(transcript_data, 'data'):
            transcript_data = transcript_data.data
        elif hasattr(transcript_data, 'transcript'):
            transcript_data = transcript_data.transcript
    
    # Convert objects to dicts
    result = []
    for item in transcript_data:
        if isinstance(item, dict):
            result.append(item)
        else:
            # Convert object to dict
            item_dict = {
                'text': getattr(item, 'text', getattr(item, 'content', '')),
                'start': getattr(item, 'start', getattr(item, 'timestamp', 0)),
                'duration': getattr(item, 'duration', getattr(item, 'dur', 3.0))
            }
            result.append(item_dict)
    
    return json.dumps(result, indent=2, ensure_ascii=False)

def format_as_srt(transcript_data):
    """Format transcript as SRT subtitle format"""
    # Convert object to list if needed
    if not isinstance(transcript_data, list):
        if hasattr(transcript_data, '__iter__'):
            transcript_data = list(transcript_data)
        elif hasattr(transcript_data, 'text') and hasattr(transcript_data, 'start'):
            transcript_data = [transcript_data]
        elif hasattr(transcript_data, 'data'):
            transcript_data = transcript_data.data
        elif hasattr(transcript_data, 'transcript'):
            transcript_data = transcript_data.transcript
    
    srt_lines = []
    for index, item in enumerate(transcript_data, start=1):
        if isinstance(item, dict):
            start = item['start']
            duration = item.get('duration', 3.0)
            text = item['text'].strip().replace('\n', ' ')
        else:
            start = getattr(item, 'start', getattr(item, 'timestamp', 0))
            duration = getattr(item, 'duration', getattr(item, 'dur', 3.0))
            text = getattr(item, 'text', getattr(item, 'content', '')).strip().replace('\n', ' ')
        
        start_time = format_time(start)
        end_time = format_time(start + duration)
        srt_lines.append(f"{index}\n{start_time} --> {end_time}\n{text}\n")
    return '\n'.join(srt_lines)

def format_as_vtt(transcript_data):
    """Format transcript as VTT subtitle format"""
    # Convert object to list if needed
    if not isinstance(transcript_data, list):
        if hasattr(transcript_data, '__iter__'):
            transcript_data = list(transcript_data)
        elif hasattr(transcript_data, 'text') and hasattr(transcript_data, 'start'):
            transcript_data = [transcript_data]
        elif hasattr(transcript_data, 'data'):
            transcript_data = transcript_data.data
        elif hasattr(transcript_data, 'transcript'):
            transcript_data = transcript_data.transcript
    
    vtt_lines = ["WEBVTT", ""]
    for item in transcript_data:
        if isinstance(item, dict):
            start = item['start']
            duration = item.get('duration', 3.0)
            text = item['text'].strip().replace('\n', ' ')
        else:
            start = getattr(item, 'start', getattr(item, 'timestamp', 0))
            duration = getattr(item, 'duration', getattr(item, 'dur', 3.0))
            text = getattr(item, 'text', getattr(item, 'content', '')).strip().replace('\n', ' ')
        
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
    
    try:
        # Get transcript
        language_list = [args.language] if args.language else ['en']
        if args.language != 'en':
            language_list.append('en')
        
        transcript_data = get_transcript(args.video_id, language_list)
        
        # Normalize transcript_data to a list format
        # Handle FetchedTranscriptSnippet or other object types
        if not isinstance(transcript_data, list):
            try:
                # Try to convert to list if it's iterable
                if hasattr(transcript_data, '__iter__'):
                    transcript_data = list(transcript_data)
                # If it has a data attribute, use that
                elif hasattr(transcript_data, 'data'):
                    transcript_data = transcript_data.data
                    if not isinstance(transcript_data, list):
                        transcript_data = list(transcript_data) if hasattr(transcript_data, '__iter__') else [transcript_data]
                # If it's a single object with text/start, wrap it in a list
                elif hasattr(transcript_data, 'text') or hasattr(transcript_data, 'start'):
                    transcript_data = [transcript_data]
                else:
                    # Last resort: try to iterate and convert
                    transcript_data = [item for item in transcript_data] if hasattr(transcript_data, '__iter__') else [transcript_data]
            except Exception:
                # If conversion fails, wrap in list
                transcript_data = [transcript_data] if transcript_data else []
        
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
        
        # Get entries count - handle both list and object
        if isinstance(transcript_data, list):
            entries_count = len(transcript_data)
        elif hasattr(transcript_data, '__len__'):
            entries_count = len(transcript_data)
        elif hasattr(transcript_data, '__iter__'):
            entries_count = len(list(transcript_data))
        else:
            entries_count = 1
        
        # Return as JSON for API
        result = {
            'success': True,
            'format': args.format,
            'content': output,
            'content_type': content_type,
            'video_id': args.video_id,
            'entries_count': entries_count
        }
        
        print(json.dumps(result))
        
    except TranscriptsDisabled:
        error_result = {
            'success': False,
            'error': 'Transcripts are disabled for this video',
            'video_id': args.video_id
        }
        print(json.dumps(error_result))
        sys.exit(1)
    except NoTranscriptFound:
        error_result = {
            'success': False,
            'error': 'No transcript found for this video in the specified language(s)',
            'video_id': args.video_id
        }
        print(json.dumps(error_result))
        sys.exit(1)
    except VideoUnavailable:
        error_result = {
            'success': False,
            'error': 'Video is unavailable or doesn\'t exist',
            'video_id': args.video_id
        }
        print(json.dumps(error_result))
        sys.exit(1)
    except Exception as e:
        error_result = {
            'success': False,
            'error': str(e),
            'video_id': args.video_id
        }
        print(json.dumps(error_result))
        sys.exit(1)

if __name__ == '__main__':
    main()
