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
    # Debug: Check what methods are available
    available_methods = [m for m in dir(YouTubeTranscriptApi) if not m.startswith('_')]
    
    # Try different API patterns that might work with version 1.2.3
    
    # Pattern 1: Check if get_transcript exists as a class method
    if 'get_transcript' in available_methods:
        try:
            if language_codes:
                for lang_code in language_codes:
                    try:
                        transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=[lang_code])
                        return transcript
                    except (NoTranscriptFound, TranscriptsDisabled):
                        continue
                    except Exception as e:
                        if 'get_transcript' in str(e):
                            break
                        continue
            
            # Try English
            try:
                transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=['en'])
                return transcript
            except (NoTranscriptFound, TranscriptsDisabled):
                pass
            except Exception:
                pass
            
            # Try without language
            try:
                transcript = YouTubeTranscriptApi.get_transcript(video_id)
                return transcript
            except (NoTranscriptFound, TranscriptsDisabled):
                pass
            except Exception:
                pass
        except Exception:
            pass
    
    # Pattern 1b: Try direct class method call anyway (in case dir() doesn't show it)
    try:
        if language_codes:
            for lang_code in language_codes:
                try:
                    transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=[lang_code])
                    return transcript
                except (NoTranscriptFound, TranscriptsDisabled):
                    continue
                except AttributeError:
                    break
                except Exception:
                    continue
        
        try:
            transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=['en'])
            return transcript
        except (NoTranscriptFound, TranscriptsDisabled, AttributeError):
            pass
        except Exception:
            pass
        
        try:
            transcript = YouTubeTranscriptApi.get_transcript(video_id)
            return transcript
        except (NoTranscriptFound, TranscriptsDisabled, AttributeError):
            pass
        except Exception:
            pass
    except Exception:
        pass
    
    # Pattern 2: Check if it's a function in the module
    try:
        import youtube_transcript_api as yt_api
        if hasattr(yt_api, 'get_transcript') and callable(getattr(yt_api, 'get_transcript')):
            if language_codes:
                for lang_code in language_codes:
                    try:
                        transcript = yt_api.get_transcript(video_id, languages=[lang_code])
                        return transcript
                    except (NoTranscriptFound, TranscriptsDisabled):
                        continue
            try:
                return yt_api.get_transcript(video_id, languages=['en'])
            except (NoTranscriptFound, TranscriptsDisabled):
                pass
            return yt_api.get_transcript(video_id)
    except Exception:
        pass
    
    # Pattern 3: Try instance method
    try:
        api_instance = YouTubeTranscriptApi()
        if hasattr(api_instance, 'get_transcript'):
            if language_codes:
                for lang_code in language_codes:
                    try:
                        transcript = api_instance.get_transcript(video_id, languages=[lang_code])
                        return transcript
                    except (NoTranscriptFound, TranscriptsDisabled):
                        continue
            try:
                return api_instance.get_transcript(video_id, languages=['en'])
            except (NoTranscriptFound, TranscriptsDisabled):
                pass
            return api_instance.get_transcript(video_id)
    except Exception:
        pass
    
    # Pattern 4: Try using list_transcripts if available (newer API)
    try:
        if hasattr(YouTubeTranscriptApi, 'list_transcripts'):
            transcript_list = YouTubeTranscriptApi.list_transcripts(video_id)
            if language_codes:
                for lang_code in language_codes:
                    try:
                        transcript = transcript_list.find_transcript([lang_code])
                        return transcript.fetch()
                    except:
                        continue
            # Try to get any transcript
            for transcript in transcript_list:
                return transcript.fetch()
    except Exception:
        pass
    
    # If nothing works, provide detailed error information
    available_methods = [m for m in dir(YouTubeTranscriptApi) if not m.startswith('_')]
    error_msg = f"Unable to fetch transcript. Available methods on YouTubeTranscriptApi: {', '.join(available_methods)}. "
    error_msg += "Please check if the video has transcripts available and if youtube-transcript-api version 1.2.3 is properly installed."
    raise Exception(error_msg)

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
        start_time = format_time(item['start'])
        end_time = format_time(item['start'] + item.get('duration', 3.0))
        text = item['text'].strip().replace('\n', ' ')
        srt_lines.append(f"{index}\n{start_time} --> {end_time}\n{text}\n")
    return '\n'.join(srt_lines)

def format_as_vtt(transcript_data):
    """Format transcript as VTT subtitle format"""
    vtt_lines = ["WEBVTT", ""]
    for item in transcript_data:
        start_time = format_time_vtt(item['start'])
        end_time = format_time_vtt(item['start'] + item.get('duration', 3.0))
        text = item['text'].strip().replace('\n', ' ')
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
        
        # Return as JSON for API
        result = {
            'success': True,
            'format': args.format,
            'content': output,
            'content_type': content_type,
            'video_id': args.video_id,
            'entries_count': len(transcript_data)
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
