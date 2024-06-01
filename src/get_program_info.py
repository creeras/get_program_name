import win32gui
import win32process
from colorama import init, Fore, Style

# Initialize colorama
init(autoreset=True)

def get_program_information():
    """활성화된 프로그램 정보를 가져옵니다."""

    # 1. 활성화된 프로그램 목록을 가져옵니다.
    programs = []
    print("Calling EnumWindows with programs:", programs)
    win32gui.EnumWindows(enum_windows_callback, programs)

    # 제목이 없는 창을 제외합니다.
    programs = [(title, pid) for title, pid in programs if title]

    # 2. 프로그램 목록을 출력하고 사용자에게 선택을 요청합니다.
    print("활성화된 프로그램 목록:")
    for idx, (title, pid) in enumerate(programs):
        print(f"{idx}: {title} (PID: {pid})")
    selected_idx = int(input("Enter the number of the program to select: "))

    # 유효하지 않은 인덱스를 입력한 경우 예외를 발생시킵니다.
    if selected_idx < 0 or selected_idx >= len(programs):
        raise Exception("Invalid selection")

    # 선택된 프로그램의 제목과 PID를 가져옵니다.
    selected_program_title, selected_program_pid = programs[selected_idx]

    # 3. 선택된 프로그램 정보를 가져옵니다.
    program_information = get_program_info(selected_program_pid)

    # 4. 프로그램 정보를 출력하고 사용자에게 선택을 요청합니다.
    print(f"\n{selected_program_title}의 프로그램 정보:")
    keys = list(program_information.keys())
    for idx, key in enumerate(keys):
        print(f"{idx}: {key} - {program_information[key]}")
    
    selected_key_idx = int(input("Enter the number of the information to return: "))

    # 유효하지 않은 인덱스를 입력한 경우 예외를 발생시킵니다.
    if selected_key_idx < 0 or selected_key_idx >= len(keys):
        raise Exception("Invalid key selection")

    selected_key = keys[selected_key_idx]
    selected_value = program_information[selected_key]

    # 선택된 정보를 반환합니다.
    result_message = f"{selected_program_title}'s {selected_key} is: {selected_value}"
    print(Fore.GREEN + Style.BRIGHT + result_message)
    return selected_value

# 콜백 함수
def enum_windows_callback(hwnd, programs):
    if win32gui.IsWindowVisible(hwnd):
        window_title = win32gui.GetWindowText(hwnd)
        _, window_pid = win32process.GetWindowThreadProcessId(hwnd)  # Get PID
        programs.append((window_title, window_pid))

# 프로그램 정보 가져오기 함수
def get_program_info(pid):
    """프로그램 정보를 가져옵니다."""
    program_info = {}
    hwnd = find_window_by_pid(pid)
    window_title = win32gui.GetWindowText(hwnd)
    window_position = win32gui.GetWindowPlacement(hwnd)
    window_size = get_window_size(hwnd)
    window_icon = get_window_icon(hwnd)
    window_process_id = pid
    program_info["title"] = window_title
    program_info["icon"] = window_icon
    program_info["position"] = window_position
    program_info["size"] = window_size
    program_info["pid"] = window_process_id
    return program_info

# 창의 크기를 가져오는 함수
def get_window_size(hwnd):
    rect = win32gui.GetWindowRect(hwnd)
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]
    return (width, height)

# 아이콘을 가져오는 함수 (임시로 문자열 반환)
def get_window_icon(hwnd):
    return "Window icon (placeholder)"

# PID로 창 찾기 함수
def find_window_by_pid(pid):
    """PID로 창을 찾습니다."""
    def callback(hwnd, pid_list):
        if win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
            pid_list.append(hwnd)
    pid_list = []
    win32gui.EnumWindows(callback, pid_list)
    if not pid_list:
        raise Exception(f"Window with PID {pid} not found")
    return pid_list[0]

# 코드 실행 예시
if __name__ == "__main__":
    program_information = get_program_information()
    print(Fore.CYAN + Style.BRIGHT + f"\nReturned information is: {program_information}")
