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
    # In version 1.2.3, the API uses 'list' and 'fetch' methods
    # Pattern 1: Try fetch() directly as a class method
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
        
        # Try English
        try:
            transcript = YouTubeTranscriptApi.fetch(video_id, languages=['en'])
            return transcript
        except (NoTranscriptFound, TranscriptsDisabled):
            pass
        except Exception:
            pass
        
        # Try without language
        try:
            transcript = YouTubeTranscriptApi.fetch(video_id)
            return transcript
        except (NoTranscriptFound, TranscriptsDisabled):
            pass
        except Exception:
            pass
    except Exception:
        pass
    
    # Pattern 2: Use list() to get available transcripts, then fetch() on each
    try:
        # Get list of available transcripts
        transcript_list = YouTubeTranscriptApi.list(video_id)
        
        # Try preferred languages first
        if language_codes:
            for lang_code in language_codes:
                try:
                    # Iterate through available transcripts
                    if hasattr(transcript_list, '__iter__'):
                        for transcript_item in transcript_list:
                            try:
                                # Check if this transcript matches our language
                                item_lang = None
                                if hasattr(transcript_item, 'language_code'):
                                    item_lang = transcript_item.language_code
                                elif hasattr(transcript_item, 'language'):
                                    item_lang = transcript_item.language
                                
                                if item_lang == lang_code:
                                    return transcript_item.fetch()
                            except Exception:
                                continue
                except (NoTranscriptFound, TranscriptsDisabled):
                    continue
                except Exception:
                    continue
        
        # Try English as fallback
        try:
            if hasattr(transcript_list, '__iter__'):
                for transcript_item in transcript_list:
                    try:
                        item_lang = None
                        if hasattr(transcript_item, 'language_code'):
                            item_lang = transcript_item.language_code
                        elif hasattr(transcript_item, 'language'):
                            item_lang = transcript_item.language
                        
                        if item_lang == 'en':
                            return transcript_item.fetch()
                    except Exception:
                        continue
        except Exception:
            pass
        
        # Try to get any available transcript
        try:
            if hasattr(transcript_list, '__iter__'):
                for transcript_item in transcript_list:
                    try:
                        return transcript_item.fetch()
                    except Exception:
                        continue
        except Exception:
            pass
        
        # If list() returns a list directly
        if isinstance(transcript_list, list) and len(transcript_list) > 0:
            return transcript_list[0].fetch()
        
        # If we get here, no transcript was found
        raise Exception("No transcript available for this video")
        
    except TranscriptsDisabled:
        raise Exception("Transcripts are disabled for this video")
    except NoTranscriptFound:
        raise Exception("No transcript found for this video")
    except VideoUnavailable:
        raise Exception("Video is unavailable or doesn't exist")
    except Exception as e:
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
